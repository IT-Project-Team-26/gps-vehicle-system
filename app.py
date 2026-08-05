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
    .status-badge {
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: bold;
        display: inline-block;
    }
    .status-working {
        background-color: #28a745;
        color: white;
    }
    .status-not-working {
        background-color: #dc3545;
        color: white;
    }
    .status-pending {
        background-color: #ffc107;
        color: black;
    }
    </style>
""", unsafe_allow_html=True)

class KPIDataProcessor:
    """Processing class for KPI data with file uploads"""
    
    def __init__(self):
        self.vm_df = None
        self.gps_df = None
        self.kpi_df = None
        self.final_df = None
    
    def load_vehicle_master(self, file):
        """Load vehicle master file"""
        try:
            df = pd.read_excel(file, engine='openpyxl', sheet_name='vehiclemaster', skiprows=4)
            df.columns = df.columns.str.strip()
            df.rename(columns={'Register Number': 'Vehicle Number'}, inplace=True)
            self.vm_df = df
            return df
        except Exception as e:
            st.error(f"Error loading vehicle master: {str(e)}")
            return None
    
    def process_kpi_file(self, file, kpi_type):
        """Process KPI files based on type"""
        try:
            df = pd.read_excel(file, engine='openpyxl')
            # FIX: Corrected the DataFrame operations
            df = df[['Kpi Date', 'Zone', 'Vehicle Number', 'Marching In Out Timings']].copy()
            df = df[df['Zone'].notna()]
            df['Kpi Source'] = kpi_type
            return df
        except Exception as e:
            st.error(f"Error processing {kpi_type}: {str(e)}")
            return None
    
    def process_kpi52(self, file):
        """Process KPI 52 with vehicle master merge"""
        try:
            df = pd.read_excel(file, engine='openpyxl')
            df.rename(columns={'Vehicle Number': 'V ID'}, inplace=True)
            merge = pd.merge(df, self.vm_df, how='left', on='V ID')
            kpi_data = merge[['Kpi Date', 'Zone_x', 'Vehicle Number', 'Marching In Out Timings']].copy()
            kpi_data = kpi_data.rename(columns={'Zone_x': 'Zone'})
            kpi_data = kpi_data[kpi_data['Zone'].notna()]
            kpi_data['Kpi Source'] = 'KPI 52'
            return kpi_data
        except Exception as e:
            st.error(f"Error processing KPI 52: {str(e)}")
            return None
    
    def process_gps_status(self, file):
        """Process GPS status file"""
        try:
            df = pd.read_excel(file, engine='openpyxl')
            df.rename(columns={'Chassis No.':'V Id'} ,inplace=True)
            
            # Age calculation
            log_dates = pd.to_datetime(df['Last Log Received At'], dayfirst=True).dt.normalize()
            max_date = log_dates.max()
            df['Age'] = (max_date - log_dates).dt.days
            
            # GPS status
            df['Status'] = np.where(df['Age'] <= 1, 'Working', 'Not Working')
            df['Date'] = max_date.strftime("%d-%m-%Y")
            
            df.rename(columns={'Vehicle Registration No.': 'Vehicle Number'}, inplace=True)
            
            final_df = df[['Date', 'GPS IMEI No.', 'Vehicle Number', 'V Id', 'Vehicle Type', 
                          'Last Log Received At', 'Last Location', 'Age', 'Status']]
            
            self.gps_df = final_df
            return final_df
        except Exception as e:
            st.error(f"Error processing GPS status: {str(e)}")
            return None
    
    def process_gps_remarks(self, file):
        """Process GPS remarks CSV file"""
        try:
            df = pd.read_csv(file, usecols=['Date', 'Vehicle Registration No.', 'Remarks', 'User', 'Lat & Long', 'Time'])
            df.rename(columns={'Vehicle Registration No.': 'Vehicle Number'}, inplace=True)
            return df
        except Exception as e:
            st.error(f"Error processing GPS remarks: {str(e)}")
            return None
    
    def combine_all_data(self, kpi_files, gps_file, remarks_file=None):
        """Combine all data sources"""
        try:
            # Process GPS
            if gps_file:
                gps_df = self.process_gps_status(gps_file)
                if gps_df is None:
                    return None
            
            # Process KPI files
            kpi_dfs = []
            for file, kpi_type in kpi_files:
                if '52' in kpi_type and self.vm_df is not None:
                    kpi_df = self.process_kpi52(file)
                else:
                    kpi_df = self.process_kpi_file(file, kpi_type)
                
                if kpi_df is not None:
                    kpi_dfs.append(kpi_df)
            
            if not kpi_dfs:
                st.warning("No valid KPI data found")
                return None
            
            # Combine all KPIs
            combined_kpi = pd.concat(kpi_dfs, ignore_index=True)
            combined_kpi = combined_kpi.drop_duplicates(subset=['Vehicle Number'], keep='first')
            
            # Merge GPS and KPI
            merge = pd.merge(gps_df, combined_kpi, how='left', on='Vehicle Number')
            
            # Filter not working vehicles
            not_working = merge[(merge['Status'] == 'Not Working') & (merge['Vehicle Number'] != 'TEST 02')]
            
            # Select columns
            not_working = not_working[['Date', 'GPS IMEI No.', 'Vehicle Number', 'V Id', 'Vehicle Type', 
                                      'Last Log Received At', 'Age', 'Status', 'Kpi Source']].copy()
            
            # Process Last Log Received At
            if 'Last Log Received At' in not_working.columns:
                not_working['Last Log Received At'] = pd.to_datetime(
                    not_working['Last Log Received At'], errors='coerce', dayfirst=True
                ).dt.strftime('%d-%m-%Y')
                
                # Handle NaN values before converting to int
                not_working['Last Log Received At'] = not_working['Last Log Received At'].fillna('01-01-1970')
                not_working['Last Log Received At'] = not_working['Last Log Received At'].astype(str)
            
            # Process remarks if available
            if remarks_file:
                remarks_df = self.process_gps_remarks(remarks_file)
                if remarks_df is not None:
                    not_working = pd.merge(not_working, remarks_df[['Vehicle Number', 'Remarks', 'User', 'Lat & Long', 'Time']], 
                                          how='left', on='Vehicle Number')
            
            # Merge with vehicle master for zone info
            if self.vm_df is not None:
                not_working = pd.merge(not_working, self.vm_df[['Vehicle Number', 'Zone', 'Facility', 'Technician']], 
                                      how='left', on='Vehicle Number')
            
            # Fill NaN values
            not_working = not_working.fillna('-')
            
            # Apply updated remarks logic
            def get_updated_remarks(row):
                remarks = str(row.get('Remarks', '')).lower() if not pd.isna(row.get('Remarks', '')) else ''
                kpi_source = row.get('Kpi Source', '')
                age = row.get('Age', 0)
                
                if remarks.startswith('wiring kit'):
                    return 'Wiring Kit Fault'
                elif remarks.startswith('converter') or remarks.startswith('dc converter'):
                    return 'Converter Fault'
                elif remarks in ('working', 'b shift', '-'):
                    return '-'
                elif remarks.startswith('line issue'):
                    return 'Line issue'
                elif remarks in ('gps issue', 'gps missing', 'gps fault', 'gps water lock') or remarks.startswith('gps'):
                    return 'GPS Fault'
                elif remarks == '':
                    return '-'
                elif kpi_source != "-" and str(remarks) not in ['Line issue', 'Gps Issue', 'Converter Fault', 'GPS Missing', 'Working', 'GPS Fault']:
                    return 'N'
                elif kpi_source == "-" and str(remarks) not in ['Line issue', 'Gps Issue', 'Converter Fault', 'GPS Missing', 'Working', 'GPS Fault']:
                    return 'BreakDown'
                else:
                    return '-'
            
            not_working['Updated Remarks'] = not_working.apply(get_updated_remarks, axis=1)
            
            self.final_df = not_working
            return not_working
            
        except Exception as e:
            st.error(f"Error combining data: {str(e)}")
            return None
    
    def get_summary_stats(self):
        """Get summary statistics from final data"""
        if self.final_df is None or self.final_df.empty:
            return {}
        
        stats = {
            'total_vehicles': len(self.final_df),
            'unique_zones': self.final_df['Zone'].nunique() if 'Zone' in self.final_df.columns else 0,
            'unique_facilities': self.final_df['Facility'].nunique() if 'Facility' in self.final_df.columns else 0,
            'total_imei': self.final_df['GPS IMEI No.'].nunique() if 'GPS IMEI No.' in self.final_df.columns else 0,
            'remarks_summary': self.final_df['Updated Remarks'].value_counts().to_dict() if 'Updated Remarks' in self.final_df.columns else {},
            'zone_summary': self.final_df['Zone'].value_counts().to_dict() if 'Zone' in self.final_df.columns else {}
        }
        
        return stats
    
    def get_technician_remarks_summary(self):
        """Get technician-wise updated remarks count"""
        if self.final_df is None or self.final_df.empty:
            return None
        
        # Create pivot table for Technician vs Updated Remarks
        tech_remarks = pd.crosstab(
            self.final_df['Technician'], 
            self.final_df['Updated Remarks'],
            margins=True,
            margins_name='Total'
        )
        
        return tech_remarks
    
    def get_visualization_remarks(self, df):
        """Replace '-' with 'Need to check' for visualization only"""
        if df is None or df.empty:
            return df
        
        # Create a copy for visualization
        viz_df = df.copy()
        
        # Replace '-' with 'Need to check' in Updated Remarks column for visualization
        if 'Updated Remarks' in viz_df.columns:
            viz_df['Updated Remarks'] = viz_df['Updated Remarks'].replace('-', 'Need to check')
        
        return viz_df


def main():
    st.markdown('<h1 class="main-header">🚛 GPS & KPI Monitoring Dashboard</h1>', unsafe_allow_html=True)
    
    # Initialize processor in session state
    if 'processor' not in st.session_state:
        st.session_state.processor = KPIDataProcessor()
        st.session_state.data_loaded = False
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 Upload Files")
        
        # Vehicle Master
        st.subheader("Vehicle Master")
        vm_file = st.file_uploader("Upload Vehicle Master", type=['xlsm', 'xlsx'])
        if vm_file:
            with st.spinner("Loading vehicle master..."):
                vm_df = st.session_state.processor.load_vehicle_master(vm_file)
                if vm_df is not None:
                    st.success(f"✅ Loaded {len(vm_df)} vehicles")
        
        # GPS Status File
        st.subheader("GPS Status")
        gps_file = st.file_uploader("Upload GPS Status", type=['xlsx', 'xlsm','csv'])
        
        # KPI Files
        st.subheader("KPI Files")
        kpi_files = []
        
        kpi_types = [
            ('KPI 52', 'kpi52'),
            ('KPI 52 Current', 'kpi52_cur'),
            ('KPI 56', 'kpi56'),
            ('KPI 63a', 'kpi63a'),
            ('KPI 63b', 'kpi63b'),
            ('KPI 72', 'kpi72')
        ]
        
        for label, key in kpi_types:
            file = st.file_uploader(f"Upload {label}", type=['xlsx'], key=key)
            if file:
                kpi_files.append((file, label))
        
        # GPS Remarks
        st.subheader("GPS Remarks")
        remarks_file = st.file_uploader("Upload Remarks", type=['csv','xlsx'])
        
        # Process Button
        st.markdown("---")
        if st.button("🚀 Process Data", type="primary", use_container_width=True):
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
        
        # Download Section
        if st.session_state.data_loaded and st.session_state.processor.final_df is not None:
            st.markdown("---")
            st.subheader("📥 Download")
            
            # Excel download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.processor.filtered_df.to_excel(writer, index=False, sheet_name='Report')
            
            excel_data = output.getvalue()
            st.download_button(
                label="📊 Download Excel Report",
                data=excel_data,
                file_name=f"gps_kpi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # Main content area
    if st.session_state.data_loaded and st.session_state.processor.final_df is not None:
        final_df = st.session_state.processor.final_df
        stats = st.session_state.processor.get_summary_stats()
        
        # Create visualization dataset with '-' replaced by 'Need to check'
        viz_df = st.session_state.processor.get_visualization_remarks(final_df)
        viz_stats = st.session_state.processor.get_summary_stats()
        # Update viz_stats with renamed remarks
        if 'Updated Remarks' in viz_df.columns:
            viz_stats['remarks_summary'] = viz_df['Updated Remarks'].value_counts().to_dict()
        
        # Summary Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("🚛 Total Vehicles", stats.get('total_vehicles', 0))
        with col2:
            st.metric("📍 Zones", stats.get('unique_zones', 0))
        with col3:
            st.metric("🏭 Facilities", stats.get('unique_facilities', 0))
        with col4:
            st.metric("📡 GPS Devices", stats.get('total_imei', 0))
        
        # Charts Section
        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📋 Data Table", "📈 Analytics"])
        
        with tab1:
            # Create two columns for charts
            col1, col2 = st.columns(2)
            
            with col1:
                # Remarks Distribution - Using visualization data
                if viz_stats.get('remarks_summary'):
                    remarks_df = pd.DataFrame({
                        'Remarks': list(viz_stats['remarks_summary'].keys()),
                        'Count': list(viz_stats['remarks_summary'].values())
                    })
                    fig = px.pie(remarks_df, values='Count', names='Remarks', title='Vehicle Status Distribution')
                    fig.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Zone Distribution
                if stats.get('zone_summary'):
                    zone_df = pd.DataFrame({
                        'Zone': list(stats['zone_summary'].keys()),
                        'Count': list(stats['zone_summary'].values())
                    })
                    fig = px.bar(zone_df, x='Zone', y='Count', title='Vehicles by Zone', 
                                color='Count', color_continuous_scale='Blues')
                    st.plotly_chart(fig, use_container_width=True)
            
            # Age Distribution
            if 'Age' in final_df.columns:
                age_df = final_df['Age'].value_counts().reset_index()
                age_df.columns = ['Days', 'Count']
                age_df = age_df.sort_values('Days')
                fig = px.line(age_df, x='Days', y='Count', title='Vehicle Age Distribution',
                             markers=True)
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            # Data Table with filters - Using original data (keeping '-')
            st.subheader("📋 Detailed Data")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            with col1:
                zone_filter = st.multiselect(
                    "Filter by Zone",
                    options=sorted(final_df['Zone'].unique()) if 'Zone' in final_df.columns else [],
                    default=[]
                )
            with col2:
                status_filter = st.multiselect(
                    "Filter by Status",
                    options=sorted(final_df['Updated Remarks'].unique()) if 'Updated Remarks' in final_df.columns else [],
                    default=[]
                )
            with col3:
                facility_filter = st.multiselect(
                    "Filter by Facility",
                    options=sorted(final_df['Facility'].unique()) if 'Facility' in final_df.columns else [],
                    default=[]
                )
            
            # Select columns for display
            display_cols = ['Date','GPS IMEI No.','Vehicle Number','V Id','Vehicle Type']
            if 'Facility' in final_df.columns:
                display_cols.append('Facility')
            display_cols.extend(['Last Log Received At','Status','Technician','Updated Remarks','Age','Kpi Source'])
            if 'Remarks' in final_df.columns:
                display_cols.extend(['Remarks','User','Lat & Long','Time'])
            
            # Apply filters
            filtered_df = final_df[display_cols].copy()
            if zone_filter and 'Zone' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Zone'].isin(zone_filter)]
            if status_filter and 'Updated Remarks' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Updated Remarks'].isin(status_filter)]
            if facility_filter and 'Facility' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Facility'].isin(facility_filter)]
            
            # Display table with original data
            st.dataframe(
                filtered_df,
                use_container_width=True,
                height=400,
                hide_index=True
            )
            
            st.caption(f"Showing {len(filtered_df)} of {len(final_df)} records")
        
        with tab3:
            st.subheader("📈 Advanced Analytics")
            
            # Vehicle Age Analysis
            col1, col2 = st.columns(2)
            
            with col1:
                # Age by Zone
                if 'Age' in final_df.columns and 'Zone' in final_df.columns:
                    age_zone = final_df.groupby('Zone')['Age'].mean().reset_index()
                    fig = px.bar(age_zone, x='Zone', y='Age', title='Average Age by Zone',
                                color='Age', color_continuous_scale='RdYlGn_r')
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Status by Zone - Using visualization data
                if 'Updated Remarks' in viz_df.columns and 'Zone' in viz_df.columns:
                    status_zone = pd.crosstab(viz_df['Zone'], viz_df['Updated Remarks'])
                    fig = px.imshow(status_zone, text_auto=True, 
                                   title='Status Distribution by Zone',
                                   color_continuous_scale='Blues')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Technician-wise Updated Remarks Analysis - Using visualization data
            st.subheader("👨‍🔧 Technician-wise Updated Remarks Analysis")
            
            if 'Technician' in viz_df.columns and 'Updated Remarks' in viz_df.columns:
                # Get the technician remarks summary using visualization data
                # Create pivot table for Technician vs Updated Remarks
                tech_remarks = pd.crosstab(
                    viz_df['Technician'], 
                    viz_df['Updated Remarks'],
                    margins=True,
                    margins_name='Total'
                )
                
                if tech_remarks is not None and not tech_remarks.empty:
                    # Display the table - show Total row in table but it's fine
                    st.dataframe(
                        tech_remarks,
                        use_container_width=True,
                        height=400
                    )
                    
                    # Visualization - Stacked bar chart
                    # Remove 'Total' row for visualization ONLY
                    tech_remarks_viz = tech_remarks.drop('Total', errors='ignore')
                    
                    # Also remove 'Total' column if it exists (margins creates both row and column totals)
                    if 'Total' in tech_remarks_viz.columns:
                        tech_remarks_viz = tech_remarks_viz.drop('Total', axis=1, errors='ignore')
                    
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
                        
                        # Additional metrics for technicians
                        st.subheader("Technician Performance Metrics")
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            # Total technicians with issues
                            # Exclude 'Need to check' column from issues count
                            issue_cols = [col for col in tech_remarks_viz.columns if col != 'Need to check']
                            if issue_cols:
                                tech_with_issues = tech_remarks_viz[issue_cols].sum(axis=1)
                                tech_with_issues = tech_with_issues[tech_with_issues > 0].count()
                            else:
                                tech_with_issues = 0
                            st.metric("Technicians with Issues", tech_with_issues)
                        
                        with col2:
                            # Most common remark per technician
                            if not tech_remarks_viz.empty:
                                # Remove 'Need to check' column if exists
                                remarks_cols = [col for col in tech_remarks_viz.columns if col != 'Need to check']
                                if remarks_cols:
                                    most_common_remarks = tech_remarks_viz[remarks_cols].idxmax(axis=1)
                                    # Get the most frequent remark across all technicians
                                    most_frequent = most_common_remarks.value_counts().index[0] if not most_common_remarks.empty else 'N/A'
                                    st.metric("Most Common Issue", most_frequent)
                                else:
                                    st.metric("Most Common Issue", "No issues")
                        
                        with col3:
                            # Total vehicles with issues
                            total_issues = 0
                            if not tech_remarks_viz.empty:
                                remarks_cols = [col for col in tech_remarks_viz.columns if col != 'Need to check']
                                if remarks_cols:
                                    total_issues = tech_remarks_viz[remarks_cols].sum().sum()
                            st.metric("Total Vehicle Issues", total_issues)
            
            # Additional Metrics
            st.subheader("Key Metrics")
            
            if 'Age' in final_df.columns:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Average Age", f"{final_df['Age'].mean():.1f} days")
                with col2:
                    st.metric("Max Age", f"{final_df['Age'].max()} days")
                with col3:
                    st.metric("Min Age", f"{final_df['Age'].min()} days")
    
    else:
        # Welcome message when no data loaded
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
        - 📋 Vehicle Master (.xlsm) - Recommended for better data matching
        
        ### Optional Files:
        - 📝 GPS Remarks (.csv) - Adds remarks and user information
        """)
        
        # Show sample workflow
        st.markdown("""
        ### Sample Workflow:
        1. Upload vehicle master to get zone, facility, and technician info
        2. Upload GPS status file to identify working/not working vehicles
        3. Upload KPI files to track performance metrics
        4. Upload remarks file for additional context
        5. Process and analyze the combined data
        """)


if __name__ == "__main__":
    main()
