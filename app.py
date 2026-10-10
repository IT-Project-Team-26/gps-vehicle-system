# app.py
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from pathlib import Path
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64

# Set page config
st.set_page_config(
    page_title="GPS & KPI Monitoring Dashboard",
    page_icon="logosumeet.jpeg",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
        border-bottom: 3px solid #1f77b4;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)


# =====================================================================
# HELPER FUNCTIONS
# =====================================================================

def normalize_vehicle(v):
    """
    Normalize vehicle numbers so KPI 5.2 and KPI 5.2 Current date
    (and other sources) all produce the same key.
    Handles: floats from Excel ('12345.0'), spaces, case, None, NaN.
    """
    if pd.isna(v):
        return ''
    s = str(v).strip().upper()
    # Remove trailing '.0' Excel adds to numeric-looking strings
    if s.endswith('.0'):
        s = s[:-2]
    # Collapse internal multiple spaces
    s = ' '.join(s.split())
    return s


def safe_sorted_unique(series):
    """Return sorted unique string values, safe against mixed dtypes."""
    if series is None:
        return []
    return sorted(pd.Series(series).dropna().astype(str).unique())


# =====================================================================
# PROCESSOR CLASS
# =====================================================================

class KPIDataProcessor:
    """Processing class for KPI data with file uploads"""

    def __init__(self):
        self.vm_df = None
        self.gps_df = None
        self.kpi_df = None
        self.final_df = None

    # -------------------------------------------------------------
    # LOADERS
    # -------------------------------------------------------------
    def load_vehicle_master(self, file):
        """Load vehicle master file"""
        try:
            df = pd.read_excel(file, engine='openpyxl',
                               sheet_name='vehiclemaster', skiprows=4)
            df.columns = df.columns.str.strip()
            df.rename(columns={'Register Number': 'Vehicle Number'}, inplace=True)
            if 'Vehicle Number' in df.columns:
                df['Vehicle Number'] = df['Vehicle Number'].apply(normalize_vehicle)
            # Dedup vehicle master on normalized key — keep first
            df = df.drop_duplicates(subset=['Vehicle Number'], keep='first')
            self.vm_df = df
            return df
        except Exception as e:
            st.error(f"Error loading vehicle master: {str(e)}")
            return None

    # -------------------------------------------------------------
    # KPI PROCESSORS
    # -------------------------------------------------------------
    def process_kpi_file(self, file, kpi_type):
        """Process generic KPI files (KPI 56, 63a, 63b, 72, etc.)"""
        try:
            df = pd.read_excel(file, engine='openpyxl')
            df.columns = df.columns.str.strip()
            df = df[['Kpi Date', 'Zone', 'Vehicle Number',
                     'Marching In Out Timings']].copy()
            df = df[df['Zone'].notna()]
            df['Vehicle Number'] = df['Vehicle Number'].apply(normalize_vehicle)
            df = df[df['Vehicle Number'] != '']
            df['Kpi Source'] = kpi_type
            # Dedup within this file
            df = df.drop_duplicates(subset=['Vehicle Number'], keep='first')
            return df
        except Exception as e:
            st.error(f"Error processing {kpi_type}: {str(e)}")
            return None

    def process_kpi52(self, file, kpi_source_label):
        """
        Process KPI 5.2 (or KPI 5.2 Current date) with vehicle master merge.
        kpi_source_label: string to put in 'Kpi Source' column.
        """
        try:
            df = pd.read_excel(file, engine='openpyxl')
            df.columns = df.columns.str.strip()

            # Detect vehicle column
            for c in ['Vehicle Number', 'V ID', 'V Id', 'Vid']:
                if c in df.columns:
                    df.rename(columns={c: 'V ID'}, inplace=True)
                    break

            # Normalize V ID in the KPI file
            if 'V ID' in df.columns:
                df['V ID'] = df['V ID'].apply(normalize_vehicle)

            # Merge with vehicle master if available
            if self.vm_df is not None and 'V ID' in df.columns:
                vm = self.vm_df.copy()
                # Normalize the merge key on VM side
                if 'Vehicle Number' in vm.columns:
                    vm['_vkey'] = vm['Vehicle Number'].apply(normalize_vehicle)
                merge = pd.merge(
                    df.assign(_vkey=df['V ID']),
                    vm, how='left', on='_vkey'
                )
            else:
                merge = df.copy()
                merge['Zone'] = '-'

            # Pick zone column
            zone_col = 'Zone_x' if 'Zone_x' in merge.columns else \
                       ('Zone' if 'Zone' in merge.columns else None)

            # Vehicle Number is either from VM or fall back to V ID
            if 'Vehicle Number' not in merge.columns:
                merge['Vehicle Number'] = merge.get('V ID', '-')

            kpi_data = merge[['Kpi Date', 'Vehicle Number',
                              'Marching In Out Timings']].copy()
            if zone_col:
                kpi_data['Zone'] = merge[zone_col]
            else:
                kpi_data['Zone'] = '-'

            kpi_data['Vehicle Number'] = kpi_data['Vehicle Number'].apply(normalize_vehicle)
            kpi_data = kpi_data[kpi_data['Vehicle Number'] != '']
            kpi_data['Kpi Source'] = kpi_source_label
            kpi_data = kpi_data.drop_duplicates(subset=['Vehicle Number'], keep='first')
            return kpi_data
        except Exception as e:
            st.error(f"Error processing {kpi_source_label}: {str(e)}")
            return None

    # -------------------------------------------------------------
    # GPS PROCESSORS
    # -------------------------------------------------------------
    def process_gps_status(self, file):
        """Process GPS status file"""
        try:
            name = getattr(file, 'name', '')
            if name.lower().endswith('.csv'):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file, engine='openpyxl')

            df.columns = df.columns.str.strip()
            df.rename(columns={'Chassis No.': 'V Id'}, inplace=True)

            # ---- Age calculation ----
            log_dates = pd.to_datetime(
                df['Last Log Received At'], dayfirst=True, errors='coerce'
            ).dt.normalize()
            max_date = log_dates.max()
            if pd.isna(max_date):
                max_date = pd.Timestamp.today().normalize()

            df['Age'] = (max_date - log_dates).dt.days
            df['Age'] = pd.to_numeric(df['Age'], errors='coerce').fillna(999).astype(int)

            # ---- GPS status ----
            df['Status'] = np.where(df['Age'] <= 1, 'Working', 'Not Working')
            df['Date'] = max_date.strftime("%d-%m-%Y")

            df.rename(columns={'Vehicle Registration No.': 'Vehicle Number'},
                      inplace=True)
            if 'Vehicle Number' in df.columns:
                df['Vehicle Number'] = df['Vehicle Number'].apply(normalize_vehicle)

            final_df = df[[
                'Date', 'GPS IMEI No.', 'Vehicle Number', 'V Id', 'Vehicle Type',
                'Last Log Received At', 'Last Location', 'Age', 'Status'
            ]].copy()

            # Dedup GPS file — one row per vehicle
            final_df = final_df.drop_duplicates(
                subset=['Vehicle Number'], keep='first'
            ).reset_index(drop=True)

            self.gps_df = final_df
            return final_df
        except Exception as e:
            st.error(f"Error processing GPS status: {str(e)}")
            return None

    def process_gps_remarks(self, file):
        """Process GPS remarks CSV/XLSX file"""
        try:
            name = getattr(file, 'name', '')
            if name.lower().endswith('.csv'):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file, engine='openpyxl')

            df.columns = df.columns.str.strip()

            for col in ['Vehicle Registration No.', 'Vehicle Number', 'Vehicle No']:
                if col in df.columns:
                    df.rename(columns={col: 'Vehicle Number'}, inplace=True)
                    break

            required = ['Vehicle Number', 'Remarks', 'User', 'Lat & Long', 'Time']
            available = [c for c in required if c in df.columns]
            df = df[available].copy()

            if 'Vehicle Number' in df.columns:
                df['Vehicle Number'] = df['Vehicle Number'].apply(normalize_vehicle)

            # Keep only the LATEST remark per vehicle to prevent duplication
            if 'Time' in df.columns and 'Date' in df.columns:
                try:
                    df['_dt'] = pd.to_datetime(
                        df['Date'].astype(str) + ' ' + df['Time'].astype(str),
                        errors='coerce', dayfirst=True
                    )
                    df = (df.sort_values('_dt', ascending=False)
                            .drop_duplicates(subset=['Vehicle Number'], keep='first')
                            .drop(columns=['_dt']))
                except Exception:
                    df = df.drop_duplicates(subset=['Vehicle Number'], keep='first')
            else:
                df = df.drop_duplicates(subset=['Vehicle Number'], keep='first')

            return df
        except Exception as e:
            st.error(f"Error processing GPS remarks: {str(e)}")
            return None

    # -------------------------------------------------------------
    # COMBINE
    # -------------------------------------------------------------
    def combine_all_data(self, kpi_files, gps_file, remarks_file=None):
        """
        Combine all data sources with priority logic.

        Priority rule (highest → lowest):
            1. KPI 5.2 Current date
            2. KPI 5.2
            3. All other KPI sources

        If a vehicle appears in multiple KPI files, keep only the
        highest-priority row.
        """
        debug_log = []

        def log(msg):
            debug_log.append(str(msg))
            print(f"[combine_all_data] {msg}")

        try:
            log("=" * 70)
            log("START combine_all_data")
            log("=" * 70)

            # ---------------------------------------------
            # STEP 1: GPS
            # ---------------------------------------------
            log("\n--- STEP 1: GPS STATUS ---")
            gps_df = self.process_gps_status(gps_file)
            if gps_df is None:
                log("GPS processing returned None — aborting")
                st.error("GPS processing failed")
                return None
            log(f"GPS rows: {len(gps_df)}")
            log(f"GPS Status counts:\n{gps_df['Status'].value_counts().to_string()}")
            log(f"GPS unique vehicles: {gps_df['Vehicle Number'].nunique()}")

            # ---------------------------------------------
            # STEP 2: KPI FILES
            # ---------------------------------------------
            log("\n--- STEP 2: KPI FILES ---")
            kpi_dfs = []
            for idx, (file, kpi_source_label) in enumerate(kpi_files, 1):
                log(f"\n  [{idx}/{len(kpi_files)}] {kpi_source_label}")

                if '5.2' in kpi_source_label or '52' in kpi_source_label:
                    kpi_df = self.process_kpi52(file, kpi_source_label)
                else:
                    kpi_df = self.process_kpi_file(file, kpi_source_label)

                if kpi_df is None or kpi_df.empty:
                    log(f"     skipped (None or empty)")
                    continue

                log(f"     rows: {len(kpi_df)}, "
                    f"unique vehicles: {kpi_df['Vehicle Number'].nunique()}, "
                    f"Kpi Source: {kpi_df['Kpi Source'].unique().tolist()}")
                kpi_dfs.append(kpi_df)

            if not kpi_dfs:
                log("No valid KPI data — aborting")
                st.warning("No valid KPI data found")
                return None

            # ---------------------------------------------
            # STEP 3: CONCAT + PRIORITY DEDUP
            # ---------------------------------------------
            log("\n--- STEP 3: CONCAT + PRIORITY DEDUP ---")
            combined_kpi = pd.concat(kpi_dfs, ignore_index=True)
            log(f"After concat: {len(combined_kpi)} rows")
            log(f"Source breakdown:\n{combined_kpi['Kpi Source'].value_counts().to_string()}")

            # Priority map — LOWER number wins
            priority_map = {
                'KPI 5.2 Current date': 0,   # highest priority
                'KPI 5.2':              1,
            }
            combined_kpi['_priority'] = (
                combined_kpi['Kpi Source'].map(priority_map).fillna(2).astype(int)
            )

            # How many vehicles appear in BOTH 5.2 and 5.2 Current date?
            v52 = set(combined_kpi.loc[
                combined_kpi['Kpi Source'] == 'KPI 5.2', 'Vehicle Number'])
            v52c = set(combined_kpi.loc[
                combined_kpi['Kpi Source'] == 'KPI 5.2 Current date', 'Vehicle Number'])
            overlap = v52 & v52c
            log(f"Vehicles in KPI 5.2               : {len(v52)}")
            log(f"Vehicles in KPI 5.2 Current date  : {len(v52c)}")
            log(f"Vehicles in BOTH (5.2 + 5.2 curr) : {len(overlap)}")
            if overlap:
                log(f"   Examples: {list(overlap)[:5]}")

            # Sort: highest priority first per vehicle, then drop dupes
            before = len(combined_kpi)
            combined_kpi = (
                combined_kpi
                .sort_values(['Vehicle Number', '_priority'],
                             ascending=[True, True])
                .drop_duplicates(subset=['Vehicle Number'], keep='first')
                .drop(columns=['_priority'])
                .reset_index(drop=True)
            )
            log(f"Dedup: {before} → {len(combined_kpi)} "
                f"(removed {before - len(combined_kpi)})")

            # Verify: no overlapping vehicle should still say 'KPI 5.2'
            wrong = combined_kpi[
                (combined_kpi['Kpi Source'] == 'KPI 5.2') &
                (combined_kpi['Vehicle Number'].isin(overlap))
            ]
            if wrong.empty:
                log("✅ Priority dedup correct — all overlapping vehicles "
                    "use KPI 5.2 Current date")
            else:
                log(f"❌ {len(wrong)} vehicle(s) still on KPI 5.2")
                log(wrong[['Vehicle Number', 'Kpi Source']].head().to_string())

            log(f"Final KPI Source breakdown:\n"
                f"{combined_kpi['Kpi Source'].value_counts().to_string()}")

            # ---------------------------------------------
            # STEP 4: MERGE GPS + KPI
            # ---------------------------------------------
            log("\n--- STEP 4: MERGE GPS + KPI ---")
            log(f"GPS rows: {len(gps_df)}  |  KPI rows: {len(combined_kpi)}")

            merge = pd.merge(gps_df, combined_kpi, how='left',
                             on='Vehicle Number')
            log(f"Merged rows: {len(merge)}")
            log(f"Rows without KPI match: {merge['Kpi Source'].isna().sum()}")

            # ---------------------------------------------
            # STEP 5: FILTER NOT WORKING
            # ---------------------------------------------
            log("\n--- STEP 5: FILTER NOT WORKING ---")
            not_working = merge[
                (merge['Status'] == 'Not Working') &
                (merge['Vehicle Number'] != 'TEST 02')
            ].copy()
            log(f"Rows after filter: {len(not_working)}")

            # ---------------------------------------------
            # STEP 6: SELECT CORE COLUMNS
            # ---------------------------------------------
            log("\n--- STEP 6: SELECT CORE COLUMNS ---")
            core_cols = ['Date', 'GPS IMEI No.', 'Vehicle Number', 'V Id',
                         'Vehicle Type', 'Last Log Received At', 'Age',
                         'Status', 'Kpi Source']
            core_cols = [c for c in core_cols if c in not_working.columns]
            not_working = not_working[core_cols].copy()
            log(f"Selected: {core_cols}")

            # ---------------------------------------------
            # STEP 7: FORMAT Last Log Received At
            # ---------------------------------------------
            if 'Last Log Received At' in not_working.columns:
                not_working['Last Log Received At'] = pd.to_datetime(
                    not_working['Last Log Received At'],
                    errors='coerce', dayfirst=True
                ).dt.strftime('%d-%m-%Y').fillna('01-01-1970')
                log("Formatted Last Log Received At")

            # ---------------------------------------------
            # STEP 8: MERGE REMARKS
            # ---------------------------------------------
            log("\n--- STEP 8: MERGE REMARKS ---")
            if remarks_file:
                remarks_df = self.process_gps_remarks(remarks_file)
                if remarks_df is not None and not remarks_df.empty:
                    before = len(not_working)
                    not_working = pd.merge(
                        not_working, remarks_df,
                        how='left', on='Vehicle Number'
                    )
                    log(f"Rows after remarks merge: {before} → {len(not_working)}")
                else:
                    log("Remarks empty/None — skipped")
            else:
                log("No remarks file")

            # ---------------------------------------------
            # STEP 9: MERGE VEHICLE MASTER
            # ---------------------------------------------
            log("\n--- STEP 9: MERGE VEHICLE MASTER ---")
            if self.vm_df is not None:
                vm_cols = ['Vehicle Number']
                for c in ['Zone', 'Facility', 'Technician']:
                    if c in self.vm_df.columns:
                        vm_cols.append(c)

                # Avoid duplicate Zone column
                if 'Zone' in not_working.columns and 'Zone' in vm_cols:
                    vm_cols = [c for c in vm_cols if c != 'Zone']

                vm_subset = (self.vm_df[vm_cols]
                             .drop_duplicates(subset=['Vehicle Number'],
                                              keep='first'))
                before = len(not_working)
                not_working = pd.merge(
                    not_working, vm_subset, how='left', on='Vehicle Number'
                )
                log(f"Rows after VM merge: {before} → {len(not_working)}")
            else:
                log("No vehicle master")

            # ---------------------------------------------
            # STEP 10: ENSURE REQUIRED COLUMNS
            # ---------------------------------------------
            required = ['Zone', 'Facility', 'Technician', 'Remarks', 'User',
                        'Lat & Long', 'Time', 'Kpi Source']
            for col in required:
                if col not in not_working.columns:
                    not_working[col] = '-'
            log("Required columns ensured")

            # ---------------------------------------------
            # STEP 11: UPDATED REMARKS LOGIC
            # ---------------------------------------------
            log("\n--- STEP 11: UPDATED REMARKS LOGIC ---")

            def get_updated_remarks(row):
                remarks_raw = row.get('Remarks', '')
                if pd.isna(remarks_raw):
                    remarks = ''
                else:
                    remarks = str(remarks_raw).strip().lower()

                kpi_source = str(row.get('Kpi Source', '-')).strip()

                # Normalize empty
                if remarks in ('nan', 'none', 'nat'):
                    remarks = ''

                if remarks.startswith('wiring kit'):
                    return 'Wiring Kit Fault'
                elif remarks.startswith('converter') or remarks.startswith('dc converter'):
                    return 'Converter Fault'
                elif remarks in ('working', 'b shift', '-'):
                    return '-'
                elif remarks.startswith('line issue'):
                    return 'Line issue'
                elif remarks in ('gps issue', 'gps missing', 'gps fault',
                                 'gps water lock') or remarks.startswith('gps'):
                    return 'GPS Fault'
                elif remarks == '':
                    return '-'
                elif kpi_source not in ('-', 'nan', 'None', '', 'NaT') and remarks not in [
                    'line issue', 'gps issue', 'converter fault',
                    'gps missing', 'working', 'gps fault'
                ]:
                    return 'N'
                elif kpi_source in ('-', 'nan', 'None', '', 'NaT'):
                    return 'BreakDown'
                else:
                    return '-'

            not_working['Updated Remarks'] = not_working.apply(
                get_updated_remarks, axis=1
            )
            log(f"Updated Remarks:\n"
                f"{not_working['Updated Remarks'].value_counts().to_string()}")

            # ---------------------------------------------
            # STEP 12: FINAL DTYPE CLEANUP
            # ---------------------------------------------
            log("\n--- STEP 12: FINAL DTYPE CLEANUP ---")

            if 'Age' in not_working.columns:
                not_working['Age'] = pd.to_numeric(
                    not_working['Age'], errors='coerce'
                ).fillna(0).astype(int)
                log(f"Age dtype: {not_working['Age'].dtype}")

            for col in not_working.columns:
                if col == 'Age':
                    continue
                not_working[col] = (
                    not_working[col]
                    .fillna('-')
                    .astype(str)
                    .replace({'nan': '-', 'None': '-', 'NaT': '-'})
                )

            # ---------------------------------------------
            # STEP 13: FINAL SAFETY DEDUP
            # ---------------------------------------------
            # Belt-and-braces: ensure no duplicate Vehicle Number survives
            before = len(not_working)
            not_working = (not_working
                           .sort_values('Kpi Source')  # ensures 5.2 current first
                           .drop_duplicates(subset=['Vehicle Number'], keep='first')
                           .reset_index(drop=True))
            if len(not_working) < before:
                log(f"⚠️ Final safety dedup removed {before - len(not_working)} "
                    f"row(s)")

            log("\n--- FINAL SUMMARY ---")
            log(f"Final rows: {len(not_working)}")
            log(f"Final Kpi Source:\n{not_working['Kpi Source'].value_counts().to_string()}")
            log(f"Final Status:\n{not_working['Status'].value_counts().to_string()}")

            self.debug_log = debug_log
            self.final_df = not_working
            return not_working

        except Exception as e:
            log(f"❌ EXCEPTION: {e}")
            import traceback
            tb = traceback.format_exc()
            log(tb)
            self.debug_log = debug_log
            st.error(f"Error combining data: {str(e)}")
            st.code(tb, language='python')
            return None

    # -------------------------------------------------------------
    # STATS & VISUALS
    # -------------------------------------------------------------
    def get_summary_stats(self):
        if self.final_df is None or self.final_df.empty:
            return {}
        df = self.final_df
        return {
            'total_vehicles': len(df),
            'unique_zones': df['Zone'].nunique() if 'Zone' in df.columns else 0,
            'unique_facilities': df['Facility'].nunique() if 'Facility' in df.columns else 0,
            'total_imei': df['GPS IMEI No.'].nunique() if 'GPS IMEI No.' in df.columns else 0,
            'remarks_summary': df['Updated Remarks'].value_counts().to_dict()
                                if 'Updated Remarks' in df.columns else {},
            'zone_summary': df['Zone'].value_counts().to_dict()
                            if 'Zone' in df.columns else {},
        }

    def get_visualization_remarks(self, df):
        if df is None or df.empty:
            return df
        viz_df = df.copy()
        if 'Updated Remarks' in viz_df.columns:
            viz_df['Updated Remarks'] = viz_df['Updated Remarks'].replace(
                '-', 'Need to check'
            )
        return viz_df


# =====================================================================
# MAIN
# =====================================================================

def main():
    st.markdown('<h1 class="main-header">🚛 GPS & KPI Monitoring Dashboard</h1>',
                unsafe_allow_html=True)

    if 'processor' not in st.session_state:
        st.session_state.processor = KPIDataProcessor()
        st.session_state.data_loaded = False

    # ---------------------- SIDEBAR ----------------------
    with st.sidebar:
        st.header("📁 Upload Files")

        st.subheader("Vehicle Master")
        vm_file = st.file_uploader("Upload Vehicle Master",
                                   type=['xlsm', 'xlsx'])
        if vm_file:
            with st.spinner("Loading vehicle master..."):
                vm_df = st.session_state.processor.load_vehicle_master(vm_file)
                if vm_df is not None:
                    st.success(f"✅ Loaded {len(vm_df)} vehicles")

        st.subheader("GPS Status")
        gps_file = st.file_uploader("Upload GPS Status",
                                    type=['xlsx', 'xlsm', 'csv'])

        st.subheader("KPI Files")
        kpi_files = []
        kpi_types = [
            ('KPI 5.2', 'kpi52'),
            ('KPI 5.2 Current date', 'kpi52_cur'),
            ('KPI 5.6', 'kpi56'),
            ('KPI 6.3a', 'kpi63a'),
            ('KPI 6.3b', 'kpi63b'),
            ('KPI 7.2', 'kpi72'),
        ]
        for label, key in kpi_types:
            file = st.file_uploader(f"Upload {label}",
                                    type=['xlsx'], key=key)
            if file:
                kpi_files.append((file, label))

        st.subheader("GPS Remarks")
        remarks_file = st.file_uploader("Upload Remarks",
                                        type=['csv', 'xlsx'])

        st.markdown("---")
        if st.button("🚀 Process Data", type="primary",
                     use_container_width=True):
            if gps_file and kpi_files:
                with st.spinner("Processing data..."):
                    result = st.session_state.processor.combine_all_data(
                        kpi_files=kpi_files,
                        gps_file=gps_file,
                        remarks_file=remarks_file
                    )
                    if result is not None:
                        st.session_state.data_loaded = True
                        st.success("✅ Data processed successfully!")
                    else:
                        st.error("❌ Failed to process data")
            else:
                st.warning("⚠️ Please upload GPS Status and at least one KPI file")

        # Debug log expander
        if (st.session_state.data_loaded
                and hasattr(st.session_state.processor, 'debug_log')):
            with st.expander("🐛 Debug Log", expanded=False):
                st.code("\n".join(st.session_state.processor.debug_log),
                        language='text')

        if (st.session_state.data_loaded
                and st.session_state.processor.final_df is not None):
            st.markdown("---")
            st.subheader("📥 Download")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.processor.final_df.to_excel(
                    writer, index=False, sheet_name='Report'
                )
            st.download_button(
                label="📊 Download Excel Report",
                data=output.getvalue(),
                file_name=f"gps_kpi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # ---------------------- MAIN CONTENT ----------------------
    if (st.session_state.data_loaded
            and st.session_state.processor.final_df is not None):

        final_df = st.session_state.processor.final_df
        stats = st.session_state.processor.get_summary_stats()

        viz_df = st.session_state.processor.get_visualization_remarks(final_df)
        viz_stats = stats.copy()
        if 'Updated Remarks' in viz_df.columns:
            viz_stats['remarks_summary'] = (
                viz_df['Updated Remarks'].value_counts().to_dict()
            )

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("🚛 Total Vehicles", stats.get('total_vehicles', 0))
        with col2:
            st.metric("📍 Zones", stats.get('unique_zones', 0))
        with col3:
            st.metric("🏭 Facilities", stats.get('unique_facilities', 0))
        with col4:
            st.metric("📡 GPS Devices", stats.get('total_imei', 0))

        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📋 Data Table", "📈 Analytics"])

        # =============== TAB 1 ===============
        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                if viz_stats.get('remarks_summary'):
                    remarks_df = pd.DataFrame({
                        'Remarks': list(viz_stats['remarks_summary'].keys()),
                        'Count': list(viz_stats['remarks_summary'].values())
                    })
                    fig = px.pie(remarks_df, values='Count', names='Remarks',
                                 title='Vehicle Status Distribution')
                    fig.update_traces(textposition='inside',
                                      textinfo='percent+label')
                    st.plotly_chart(fig, use_container_width=True)
            with col2:
                if stats.get('zone_summary'):
                    zone_df = pd.DataFrame({
                        'Zone': list(stats['zone_summary'].keys()),
                        'Count': list(stats['zone_summary'].values())
                    })
                    fig = px.bar(zone_df, x='Zone', y='Count',
                                 title='Vehicles by Zone',
                                 color='Count',
                                 color_continuous_scale='Blues')
                    st.plotly_chart(fig, use_container_width=True)

            if 'Age' in final_df.columns:
                age_numeric = pd.to_numeric(final_df['Age'], errors='coerce').dropna()
                if not age_numeric.empty:
                    age_df = age_numeric.value_counts().reset_index()
                    age_df.columns = ['Days', 'Count']
                    age_df = age_df.sort_values('Days')
                    fig = px.line(age_df, x='Days', y='Count',
                                  title='Vehicle Age Distribution',
                                  markers=True)
                    st.plotly_chart(fig, use_container_width=True)

        # =============== TAB 2 ===============
        with tab2:
            st.subheader("📋 Detailed Data")

            col1, col2, col3 = st.columns(3)
            with col1:
                zone_opts = safe_sorted_unique(final_df['Zone']) \
                    if 'Zone' in final_df.columns else []
                zone_filter = st.multiselect("Filter by Zone",
                                             options=zone_opts, default=[])
            with col2:
                status_opts = safe_sorted_unique(final_df['Updated Remarks']) \
                    if 'Updated Remarks' in final_df.columns else []
                status_filter = st.multiselect("Filter by Status",
                                               options=status_opts, default=[])
            with col3:
                facility_opts = safe_sorted_unique(final_df['Facility']) \
                    if 'Facility' in final_df.columns else []
                facility_filter = st.multiselect("Filter by Facility",
                                                 options=facility_opts, default=[])

            display_cols = ['Date', 'GPS IMEI No.', 'Vehicle Number', 'V Id',
                            'Vehicle Type']
            if 'Zone' in final_df.columns:
                display_cols.append('Zone')
            if 'Facility' in final_df.columns:
                display_cols.append('Facility')
            display_cols.extend(['Last Log Received At', 'Status', 'Technician',
                                 'Updated Remarks', 'Age', 'Kpi Source'])
            if 'Remarks' in final_df.columns:
                display_cols.extend(['Remarks', 'User', 'Lat & Long', 'Time'])

            display_cols = [c for c in display_cols if c in final_df.columns]
            filtered_df = final_df[display_cols].copy()

            if zone_filter and 'Zone' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Zone'].isin(zone_filter)]
            if status_filter and 'Updated Remarks' in filtered_df.columns:
                filtered_df = filtered_df[
                    filtered_df['Updated Remarks'].isin(status_filter)
                ]
            if facility_filter and 'Facility' in filtered_df.columns:
                filtered_df = filtered_df[
                    filtered_df['Facility'].isin(facility_filter)
                ]

            st.dataframe(filtered_df, use_container_width=True, hide_index=True)
            st.caption(f"Showing {len(filtered_df)} of {len(final_df)} records")

        # =============== TAB 3 ===============
        with tab3:
            st.subheader("📈 Advanced Analytics")
            col1, col2 = st.columns(2)

            with col1:
                if 'Age' in final_df.columns and 'Zone' in final_df.columns:
                    tmp = final_df.copy()
                    tmp['Age'] = pd.to_numeric(tmp['Age'], errors='coerce')
                    age_zone = tmp.groupby('Zone')['Age'].mean().reset_index()
                    fig = px.bar(age_zone, x='Zone', y='Age',
                                 title='Average Age by Zone',
                                 color='Age',
                                 color_continuous_scale='RdYlGn_r')
                    st.plotly_chart(fig, use_container_width=True)

            with col2:
                if ('Updated Remarks' in viz_df.columns
                        and 'Zone' in viz_df.columns):
                    status_zone = pd.crosstab(viz_df['Zone'],
                                              viz_df['Updated Remarks'])
                    if not status_zone.empty:
                        fig = px.imshow(status_zone, text_auto=True,
                                        title='Status Distribution by Zone',
                                        color_continuous_scale='Blues')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

            st.subheader("👨‍🔧 Technician-wise Updated Remarks Analysis")
            if ('Technician' in viz_df.columns
                    and 'Updated Remarks' in viz_df.columns):
                tech_remarks = pd.crosstab(
                    viz_df['Technician'],
                    viz_df['Updated Remarks'],
                    margins=True,
                    margins_name='Total'
                )
                if not tech_remarks.empty:
                    st.dataframe(tech_remarks, use_container_width=True,
                                 height=400)

                    tech_remarks_viz = tech_remarks.drop('Total', errors='ignore')
                    if 'Total' in tech_remarks_viz.columns:
                        tech_remarks_viz = tech_remarks_viz.drop(
                            'Total', axis=1, errors='ignore')

                    if not tech_remarks_viz.empty:
                        fig = px.bar(
                            tech_remarks_viz,
                            title='Technician-wise Updated Remarks Distribution',
                            labels={'value': 'Count', 'variable': 'Remarks'},
                            barmode='stack',
                            color_discrete_sequence=px.colors.qualitative.Set3
                        )
                        fig.update_layout(
                            xaxis_title='Technician',
                            yaxis_title='Number of Vehicles',
                            legend_title='Updated Remarks'
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        st.subheader("Technician Performance Metrics")
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            issue_cols = [c for c in tech_remarks_viz.columns
                                          if c != 'Need to check']
                            tech_with_issues = 0
                            if issue_cols:
                                t = tech_remarks_viz[issue_cols].sum(axis=1)
                                tech_with_issues = t[t > 0].count()
                            st.metric("Technicians with Issues", tech_with_issues)
                        with c2:
                            remarks_cols = [c for c in tech_remarks_viz.columns
                                            if c != 'Need to check']
                            if remarks_cols:
                                most_common = tech_remarks_viz[remarks_cols].idxmax(axis=1)
                                most_frequent = (
                                    most_common.value_counts().index[0]
                                    if not most_common.empty else 'N/A'
                                )
                                st.metric("Most Common Issue", most_frequent)
                            else:
                                st.metric("Most Common Issue", "No issues")
                        with c3:
                            total_issues = 0
                            remarks_cols = [c for c in tech_remarks_viz.columns
                                            if c != 'Need to check']
                            if remarks_cols:
                                total_issues = tech_remarks_viz[remarks_cols].sum().sum()
                            st.metric("Total Vehicle Issues", total_issues)

            st.subheader("Key Metrics")
            if 'Age' in final_df.columns:
                age_series = pd.to_numeric(final_df['Age'], errors='coerce')
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Average Age",
                              f"{age_series.mean():.1f} days"
                              if age_series.notna().any() else "N/A")
                with c2:
                    st.metric("Max Age",
                              f"{age_series.max():.0f} days"
                              if age_series.notna().any() else "N/A")
                with c3:
                    st.metric("Min Age",
                              f"{age_series.min():.0f} days"
                              if age_series.notna().any() else "N/A")

    else:
        st.info("""
        👋 **Welcome to the GPS & KPI Monitoring Dashboard!**

        ### How to use:
        1. **Upload Files** in the sidebar
        2. Click **Process Data** to generate the report
        3. View **Dashboards**, **Data Tables**, and **Analytics**
        4. **Download** the Excel report

        ### Required Files:
        - 📍 GPS Status File (.xlsx or .xlsm)
        - 📊 At least one KPI file (.xlsx)
        - 📋 Vehicle Master (.xlsm) - Recommended

        ### Optional Files:
        - 📝 GPS Remarks (.csv) - Adds remarks and user information
        """)


if __name__ == "__main__":
    main()