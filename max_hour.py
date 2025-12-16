import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Maxhour Flight Hours Analyzer",
    page_icon="✈️",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f9ff;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3b82f6;
    }
    .stDownloadButton button {
        background-color: #10b981;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.markdown("""
<div class="main-header">
    <h1 style="color: white;">✈️ Maxhour Flight Hours Analyzer</h1>
    <p>Analyze Crew With Over Flight Hours</p>
</div>
""", unsafe_allow_html=True)

# Helper functions
def decimal_flight_hours(hours):
    """Convert HH:MM format to decimal hours"""
    try:
        # Handle if already a number
        if isinstance(hours, (int, float)):
            return float(hours)
        
        # Convert to string and split by ':'
        hours_str = str(hours).strip()
        
        # If contains ':', it's in HH:MM format
        if ':' in hours_str:
            pemisah = hours_str.split(':')
            jam = int(pemisah[0])
            menit = int(pemisah[1])
            jam_desimal = jam + (menit / 60)
            return round(jam_desimal, 2)
        else:
            # If no ':', treat as decimal
            return float(hours_str.replace(',', ''))
    except:
        return 0.0

def find_column(df, possible_names):
    """Find column name from possible variations (case-insensitive)"""
    df_cols = {col: col for col in df.columns}
    df_cols_lower = {col.lower(): col for col in df.columns}
    
    for name in possible_names:
        # Exact match first
        if name in df_cols:
            return name
        # Case-insensitive match
        if name.lower() in df_cols_lower:
            return df_cols_lower[name.lower()]
    return None

def actual_rank(rank):
    """Determine actual rank (COCKPIT or CABIN)"""
    rank_str = str(rank).strip().upper()
    if rank_str in ["CPT", "FO", "CPT/FO"]:
        return "COCKPIT"
    else:
        return "CABIN"

def crew_hour_status_mon(hours):
    """Determine crew hour status for monthly (>110)"""
    if hours > 110:
        return "OVER"
    else:
        return "OTHER"

def crew_hour_status_year(hours):
    """Determine crew hour status for yearly (>1050)"""
    if hours > 1050:
        return "OVER"
    else:
        return "OTHER"

def process_monthly_data(df):
    """Process Monthly Report Flight Hours data"""
    # Find columns
    crew_id_col = find_column(df, ['Crew ID', 'ID', 'crew id', 'id'])
    fh_col = find_column(df, ['Flight Hours', 'FLIGHT HOURS', 'flight hours'])
    rank_col = find_column(df, ['Rank', 'RANK', 'rank'])
    company_col = find_column(df, ['Company', 'COMPANY', 'company'])
    crew_cat_col = find_column(df, ['Crew Category', 'CREW CATEGORY', 'crew category'])
    crew_status_col = find_column(df, ['Crew Status', 'CREW STATUS', 'crew status'])
    
    if not fh_col:
        st.error("Column 'Flight Hours' not found in Monthly Report!")
        return None
    
    # Convert Flight Hours to decimal (HH:MM to decimal)
    df['Flight Hours Decimal'] = df[fh_col].apply(decimal_flight_hours)
    
    # Add Actual Rank column
    if rank_col:
        df['Actual Rank'] = df[rank_col].apply(actual_rank)
    else:
        df['Actual Rank'] = 'CABIN'
    
    # Add Crew Hour Status column
    df['Crew Hour Status'] = df['Flight Hours Decimal'].apply(crew_hour_status_mon)
    
    return df

def process_consecutive_data(df, df_mon_standardized):
    """Process Crew Consecutive Year Flight Hours data and merge with monthly"""
    # Find columns in consecutive year data
    id_col = find_column(df, ['ID', 'Crew ID', 'id', 'crew id'])
    fh_col = find_column(df, ['FLIGHT HOURS', 'Flight Hours', 'flight hours'])
    rank_col = find_column(df, ['RANK', 'Rank', 'rank'])
    company_col = find_column(df, ['COMPANY', 'Company', 'company'])
    
    if not fh_col:
        st.error("Column 'Flight Hours' not found in Consecutive Year Report!")
        return None
    
    # Convert Flight Hours to decimal (HH:MM to decimal)
    df['Flight Hours Decimal'] = df[fh_col].apply(decimal_flight_hours)
    
    # Add Actual Rank column
    if rank_col:
        df['Actual Rank'] = df[rank_col].apply(actual_rank)
    else:
        df['Actual Rank'] = 'CABIN'
    
    # Add Crew Hour Status column
    df['Crew Hour Status'] = df['Flight Hours Decimal'].apply(crew_hour_status_year)
    
    # Merge with monthly to get Crew Category and Crew Status
    if df_mon_standardized is not None:
        crew_id_monthly = find_column(df_mon_standardized, ['Crew ID', 'ID', 'crew id', 'id'])
        crew_cat_monthly = find_column(df_mon_standardized, ['Crew Category', 'CREW CATEGORY', 'crew category'])
        crew_status_monthly = find_column(df_mon_standardized, ['Crew Status', 'CREW STATUS', 'crew status'])
        
        if crew_id_monthly and crew_cat_monthly and crew_status_monthly and id_col:
            # Merge using left join
            df_merged = df.merge(
                df_mon_standardized[[crew_id_monthly, crew_cat_monthly, crew_status_monthly]],
                how='left',
                left_on=id_col,
                right_on=crew_id_monthly
            )
            
            # Drop duplicate Crew ID column if exists
            if crew_id_monthly in df_merged.columns and crew_id_monthly != id_col:
                df_merged = df_merged.drop(columns=[crew_id_monthly])
            
            # Fill NaN with '-'
            df_merged[crew_cat_monthly] = df_merged[crew_cat_monthly].fillna('-')
            df_merged[crew_status_monthly] = df_merged[crew_status_monthly].fillna('-')
            
            # Rename to standard names
            df_merged = df_merged.rename(columns={
                crew_cat_monthly: 'Crew Category',
                crew_status_monthly: 'Crew Status'
            })
            
            return df_merged
        else:
            df['Crew Category'] = '-'
            df['Crew Status'] = '-'
            return df
    else:
        df['Crew Category'] = '-'
        df['Crew Status'] = '-'
        return df

def calculate_summary(df, is_monthly=True):
    """Calculate summary statistics per company"""
    company_col = find_column(df, ['Company', 'COMPANY', 'company'])
    id_col = find_column(df, ['Crew ID', 'ID', 'crew id', 'id'])
    
    if not company_col:
        st.warning("Column 'Company' not found!")
        return pd.DataFrame()
    
    companies = df[company_col].unique()
    summary_data = []
    
    for company in companies:
        if pd.isna(company) or company == '':
            continue
            
        company_df = df[df[company_col] == company]
        
        # Penyebut: Crew Status == Ready Crew & Actual Rank == COCKPIT
        ready_cockpit = company_df[
            (company_df['Crew Status'] == 'Ready Crew') & 
            (company_df['Actual Rank'] == 'COCKPIT')
        ]
        
        # Pembilang: Ready Crew & COCKPIT & OVER
        ready_cockpit_over = ready_cockpit[ready_cockpit['Crew Hour Status'] == 'OVER']
        
        total = len(ready_cockpit)
        over = len(ready_cockpit_over)
        percentage = (over / total * 100) if total > 0 else 0
        
        summary_data.append({
            'Company': company,
            'Total Ready Cockpit': total,
            'Over Limit': over,
            'Percentage': round(percentage, 2)
        })
    
    return pd.DataFrame(summary_data)

def create_summary_report(monthly_summary, consecutive_summary):
    """Create final summary report matching the format"""
    if monthly_summary.empty and consecutive_summary.empty:
        return pd.DataFrame()
    
    companies = sorted(list(set(
        monthly_summary['Company'].tolist() + consecutive_summary['Company'].tolist()
    )))
    
    report_data = []
    for company in companies:
        monthly_pct = monthly_summary[monthly_summary['Company'] == company]['Percentage'].values
        consecutive_pct = consecutive_summary[consecutive_summary['Company'] == company]['Percentage'].values
        
        monthly_val = f"{monthly_pct[0]:.2f}%" if len(monthly_pct) > 0 else "0.00%"
        consecutive_val = f"{consecutive_pct[0]:.2f}%" if len(consecutive_pct) > 0 else "0.00%"
        
        report_data.append({
            'Company': company,
            'Period': 'Monthly',
            'Percentage': monthly_val
        })
        report_data.append({
            'Company': '',
            'Period': '12 Consecutive Months',
            'Percentage': consecutive_val
        })
    
    return pd.DataFrame(report_data)

def export_to_excel(summary_report, monthly_over, consecutive_over, monthly_summary, consecutive_summary):
    """Export results to Excel file"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet (formatted)
        summary_report.to_excel(writer, sheet_name='Summary Report', index=False)
        
        # Detailed summaries
        monthly_summary.to_excel(writer, sheet_name='Monthly Summary', index=False)
        consecutive_summary.to_excel(writer, sheet_name='Consecutive Summary', index=False)
        
        # Monthly Over sheet
        monthly_over.to_excel(writer, sheet_name='Monthly Over', index=False)
        
        # Consecutive Over sheet
        consecutive_over.to_excel(writer, sheet_name='Consecutive Over', index=False)
    
    output.seek(0)
    return output

# Main App
st.markdown("---")

# File upload section
col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 Monthly Report Flight Hours")
    st.caption("Sheet: Standardized_Company | Header: Row 1")
    monthly_file = st.file_uploader(
        "Upload Monthly Report",
        type=['xlsx', 'xls'],
        key='monthly',
        help="File dengan sheet 'Standardized_Company'"
    )

with col2:
    st.subheader("📅 Crew Consecutive Year Flight Hours")
    st.caption("Header on Row 2")
    consecutive_file = st.file_uploader(
        "Upload Consecutive Year Report",
        type=['xlsx', 'xls'],
        key='consecutive',
        help="File dengan header di baris 2 (skip row 1)"
    )

st.markdown("---")

# Process button
if st.button("🚀 Analyze Data", type="primary", use_container_width=True):
    if monthly_file is None or consecutive_file is None:
        st.error("⚠️ Please upload both files before analyzing!")
    else:
        with st.spinner("Processing data..."):
            try:
                # Read Excel files
                df_mon_standardized = pd.read_excel(monthly_file, sheet_name='Standardized_Company')
                df_year = pd.read_excel(consecutive_file, header=1)
                
                st.info(f"📄 Monthly Report: {len(df_mon_standardized)} rows loaded")
                st.info(f"📄 Consecutive Year: {len(df_year)} rows loaded")
                
                # Process monthly data
                df_mon_processed = process_monthly_data(df_mon_standardized)
                
                if df_mon_processed is None:
                    st.stop()
                
                # Process consecutive year data with merge
                df_year_merged = process_consecutive_data(df_year, df_mon_standardized)
                
                if df_year_merged is None:
                    st.stop()
                
                # Calculate summaries
                monthly_summary = calculate_summary(df_mon_processed, is_monthly=True)
                consecutive_summary = calculate_summary(df_year_merged, is_monthly=False)
                
                # Create final report
                summary_report = create_summary_report(monthly_summary, consecutive_summary)
                
                # Get over limit crews
                monthly_over = df_mon_processed[df_mon_processed['Crew Hour Status'] == 'OVER']
                consecutive_over = df_year_merged[df_year_merged['Crew Hour Status'] == 'OVER']
                
                # Store in session state
                st.session_state['results'] = {
                    'summary_report': summary_report,
                    'monthly_summary': monthly_summary,
                    'consecutive_summary': consecutive_summary,
                    'monthly_over': monthly_over,
                    'consecutive_over': consecutive_over,
                    'monthly_processed': df_mon_processed,
                    'consecutive_processed': df_year_merged
                }
                
                st.success("✅ Data processed successfully!")
                st.balloons()
                
            except Exception as e:
                st.error(f"❌ Error processing data: {str(e)}")
                st.exception(e)

# Display results if available
if 'results' in st.session_state:
    results = st.session_state['results']
    
    st.markdown("---")
    st.header("📈 Rate of Maxhour Report")
    
    # Display summary table with better styling
    st.markdown("### 📋 Summary by Company")
    st.dataframe(
        results['summary_report'],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Company": st.column_config.TextColumn("Company", width="medium"),
            "Period": st.column_config.TextColumn("Period", width="large"),
            "Percentage": st.column_config.TextColumn("Percentage", width="small")
        }
    )
    
    # Show detailed summary
    with st.expander("📊 View Detailed Summary"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Monthly Summary**")
            st.dataframe(results['monthly_summary'], use_container_width=True, hide_index=True)
        with col2:
            st.markdown("**Consecutive Summary**")
            st.dataframe(results['consecutive_summary'], use_container_width=True, hide_index=True)
    
    # Export button
    excel_file = export_to_excel(
        results['summary_report'],
        results['monthly_over'],
        results['consecutive_over'],
        results['monthly_summary'],
        results['consecutive_summary']
    )
    
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_file,
        file_name="Maxhour_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("---")
    
    # Visualizations
    st.header("📊 Visual Analysis")
    
    # Create bar charts
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Monthly Analysis")
        if not results['monthly_summary'].empty:
            fig1 = px.bar(
                results['monthly_summary'],
                x='Company',
                y='Percentage',
                title='Monthly Maxhour Rate by Company (>110 hours)',
                labels={'Percentage': 'Percentage (%)'},
                color='Percentage',
                color_continuous_scale='Reds',
                text='Percentage'
            )
            fig1.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
            fig1.update_layout(showlegend=False, yaxis_title="Percentage (%)")
            st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        st.subheader("12 Consecutive Months Analysis")
        if not results['consecutive_summary'].empty:
            fig2 = px.bar(
                results['consecutive_summary'],
                x='Company',
                y='Percentage',
                title='Consecutive Maxhour Rate by Company (>1050 hours)',
                labels={'Percentage': 'Percentage (%)'},
                color='Percentage',
                color_continuous_scale='Reds',
                text='Percentage'
            )
            fig2.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
            fig2.update_layout(showlegend=False, yaxis_title="Percentage (%)")
            st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("---")
    
    # Crew over limit details
    st.header("👥 Crew Over Limit Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🔴 Monthly Over Limit (> 110 hours)")
        st.metric("Total Crew Over Limit", len(results['monthly_over']))
        
        if not results['monthly_over'].empty:
            # Select columns to display
            display_cols = []
            for possible_col in ['Name', 'name', 'Crew Name', 'crew name']:
                col = find_column(results['monthly_over'], [possible_col])
                if col:
                    display_cols.append(col)
                    break
            
            for possible_col in ['Company', 'company', 'COMPANY']:
                col = find_column(results['monthly_over'], [possible_col])
                if col and col not in display_cols:
                    display_cols.append(col)
                    break
            
            for possible_col in ['Rank', 'rank', 'RANK']:
                col = find_column(results['monthly_over'], [possible_col])
                if col and col not in display_cols:
                    display_cols.append(col)
                    break
            
            if 'Flight Hours Decimal' in results['monthly_over'].columns:
                display_cols.append('Flight Hours Decimal')
            
            if 'Crew Status' in results['monthly_over'].columns:
                display_cols.append('Crew Status')
            
            st.dataframe(
                results['monthly_over'][display_cols].head(20),
                use_container_width=True,
                hide_index=True
            )
            if len(results['monthly_over']) > 20:
                st.caption(f"Showing 20 of {len(results['monthly_over'])} crew")
    
    with col2:
        st.subheader("🔴 Consecutive Over Limit (> 1050 hours)")
        st.metric("Total Crew Over Limit", len(results['consecutive_over']))
        
        if not results['consecutive_over'].empty:
            # Select columns to display
            display_cols = []
            for possible_col in ['Crew Name', 'crew name', 'Name', 'name']:
                col = find_column(results['consecutive_over'], [possible_col])
                if col:
                    display_cols.append(col)
                    break
            
            for possible_col in ['Company', 'company', 'COMPANY']:
                col = find_column(results['consecutive_over'], [possible_col])
                if col and col not in display_cols:
                    display_cols.append(col)
                    break
            
            for possible_col in ['Rank', 'rank', 'RANK']:
                col = find_column(results['consecutive_over'], [possible_col])
                if col and col not in display_cols:
                    display_cols.append(col)
                    break
            
            if 'Flight Hours Decimal' in results['consecutive_over'].columns:
                display_cols.append('Flight Hours Decimal')
            
            if 'Crew Status' in results['consecutive_over'].columns:
                display_cols.append('Crew Status')
            
            st.dataframe(
                results['consecutive_over'][display_cols].head(20),
                use_container_width=True,
                hide_index=True
            )
            if len(results['consecutive_over']) > 20:
                st.caption(f"Showing 20 of {len(results['consecutive_over'])} crew")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>Maxhour Flight Hours Analyzer | © Maurino Audrian Putra</p>",
    unsafe_allow_html=True
)