import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from io import BytesIO
import xlsxwriter

# Set page config
st.set_page_config(
    page_title="Service Analytics Dashboard",
    page_icon="📚",
    layout="wide"
)

# Custom CSS (same as before)
st.markdown("""
<style>
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 0.5rem;
        color: white;
        margin-bottom: 2rem;
    }
    .kpi-container {
        display: flex;
        justify-content: center;
        gap: 20px;
        margin: 20px 0;
        flex-wrap: wrap;
    }
    .kpi-box {
        min-width: 220px;
        max-width: 220px;
        height: 130px;
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        color: white;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    .kpi-title { font-size: 14px; opacity: 0.9; margin-bottom: 8px; font-weight: 500; }
    .kpi-value { font-size: 34px; font-weight: 700; margin: 5px 0; }
    .kpi-sub { font-size: 13px; opacity: 0.85; }
    .kpi-box-blue { background: linear-gradient(135deg, #4A6FA5, #166088); }
    .kpi-box-green { background: linear-gradient(135deg, #2E8B57, #3CB371); }
    .kpi-box-orange { background: linear-gradient(135deg, #FF7A59, #FFA500); }
    .kpi-box-purple { background: linear-gradient(135deg, #8A2BE2, #9370DB); }
    .section-header {
        background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1.5rem 0 1rem 0;
        border-left: 4px solid #4dabf7;
    }
</style>
""", unsafe_allow_html=True)

def render_kpi(title, value, subtitle="", color_class=""):
    return f"""
    <div class="kpi-box {color_class}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtitle}</div>
    </div>
    """

def render_kpi_row(kpis):
    return f'<div class="kpi-container">{"".join(kpis)}</div>'

def load_csv_file(uploaded_file):
    """Load data from uploaded CSV file."""
    try:
        df = pd.read_csv(uploaded_file)
        return df, None
    except Exception as e:
        return None, str(e)

def normalize_dataframe(df):
    """Normalize column names and data."""
    if df.empty:
        return df
    
    # Standardize column names
    column_mapping = {
        'SL NO': 'SL_NO',
        'STUDENT NAME': 'Student_Name',
        'COURSE': 'Course',
        'CONTACT NUMBER': 'Contact',
        'PACKAGE (HOURS)': 'Package_Hours',
        'MAIL ID': 'Email',
        'JOINING DATE': 'Joining_Date',
        'INDIVIDUAL - STARTED / NOT STARTED': 'Status',
        ' STARTED DATE': 'Started_Date',
        'COURSE COMPLETED DATE': 'Completed_Date',
        'TUTOR NAME': 'Tutor',
        'TEAM NAME': 'Team'
    }
    
    # Rename columns if they exist
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns:
            df = df.rename(columns={old_name: new_name})
    
    # Normalize Status
    if 'Status' in df.columns:
        df['Status_Clean'] = df['Status'].astype(str).str.lower().apply(
            lambda x: 'Started' if 'start' in x or 'yes' in x else 'Not Started'
        )
    
    # Add completion status
    if 'Completed_Date' in df.columns:
        df['Is_Completed'] = df['Completed_Date'].notna() & (df['Completed_Date'] != '')
        df['Completion_Status'] = df['Is_Completed'].apply(
            lambda x: 'Completed' if x else 'In Progress'
        )
    
    return df

def calculate_kpis(df):
    """Calculate KPIs."""
    if df.empty:
        return {}
    
    total_students = len(df)
    started = len(df[df['Status_Clean'] == 'Started']) if 'Status_Clean' in df.columns else 0
    not_started = total_students - started
    completed = df['Is_Completed'].sum() if 'Is_Completed' in df.columns else 0
    in_progress = started - completed
    start_rate = round((started / total_students * 100), 1) if total_students > 0 else 0
    completion_rate = round((completed / total_students * 100), 1) if total_students > 0 else 0
    
    top_tutor = df['Tutor'].value_counts().index[0] if 'Tutor' in df.columns and not df['Tutor'].empty else 'N/A'
    top_team = df['Team'].value_counts().index[0] if 'Team' in df.columns and not df['Team'].empty else 'N/A'
    top_course = df['Course'].value_counts().index[0] if 'Course' in df.columns and not df['Course'].empty else 'N/A'
    
    return {
        'total_students': total_students,
        'started': started,
        'not_started': not_started,
        'completed': completed,
        'in_progress': in_progress,
        'start_rate': start_rate,
        'completion_rate': completion_rate,
        'top_tutor': str(top_tutor)[:20],
        'top_team': str(top_team)[:20],
        'top_course': str(top_course)[:20]
    }

def create_tutor_performance(df):
    """Create tutor performance metrics."""
    if df.empty or 'Tutor' not in df.columns:
        return pd.DataFrame()
    
    tutor_stats = df.groupby('Tutor').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', 'sum'),
        In_Progress=('Completion_Status', lambda x: (x == 'In Progress').sum())
    ).reset_index()
    
    tutor_stats['Start_Rate_%'] = np.where(
        tutor_stats['Total_Students'] > 0,
        (tutor_stats['Started'] / tutor_stats['Total_Students'] * 100).round(1),
        0
    )
    
    tutor_stats['Completion_Rate_%'] = np.where(
        tutor_stats['Started'] > 0,
        (tutor_stats['Completed'] / tutor_stats['Started'] * 100).round(1),
        0
    )
    
    return tutor_stats.sort_values('Total_Students', ascending=False)

def create_team_performance(df):
    """Create team performance metrics."""
    if df.empty or 'Team' not in df.columns:
        return pd.DataFrame()
    
    team_stats = df.groupby('Team').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', 'sum')
    ).reset_index()
    
    team_stats['Completion_Rate_%'] = np.where(
        team_stats['Started'] > 0,
        (team_stats['Completed'] / team_stats['Started'] * 100).round(1),
        0
    )
    
    return team_stats.sort_values('Total_Students', ascending=False)

def create_course_analysis(df):
    """Create course-wise analysis."""
    if df.empty or 'Course' not in df.columns:
        return pd.DataFrame()
    
    course_stats = df.groupby('Course').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', 'sum')
    ).reset_index()
    
    course_stats['Completion_Rate_%'] = np.where(
        course_stats['Started'] > 0,
        (course_stats['Completed'] / course_stats['Started'] * 100).round(1),
        0
    )
    
    return course_stats.sort_values('Total_Students', ascending=False)

def create_excel_report(df, kpis, tutor_perf, team_perf, course_analysis):
    """Create Excel report."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'align': 'center',
            'border': 1
        })
        
        # Executive Summary
        summary_data = {
            'Metric': ['Total Students', 'Started', 'Not Started', 'Completed', 'In Progress', 'Start Rate %', 'Completion Rate %'],
            'Value': [kpis['total_students'], kpis['started'], kpis['not_started'], kpis['completed'], kpis['in_progress'], kpis['start_rate'], kpis['completion_rate']]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        worksheet = writer.sheets['Executive Summary']
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        if not df.empty:
            df.to_excel(writer, sheet_name='Raw Data', index=False)
        if not tutor_perf.empty:
            tutor_perf.to_excel(writer, sheet_name='Tutor Performance', index=False)
        if not team_perf.empty:
            team_perf.to_excel(writer, sheet_name='Team Performance', index=False)
        if not course_analysis.empty:
            course_analysis.to_excel(writer, sheet_name='Course Analysis', index=False)
    
    return output.getvalue()

def main():
    st.markdown("""
        <div class="header-container">
            <h1 style="margin: 0; font-size: 2.5rem;">📚 Service Analytics Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem;">Student Progress & Performance Tracking</p>
        </div>
    """, unsafe_allow_html=True)
    
    if 'data_df' not in st.session_state:
        st.session_state.data_df = None
    
    # Sidebar
    with st.sidebar:
        st.markdown("## 📁 Upload Data")
        
        st.info("""
        **How to export from Google Sheets:**
        1. Open your Google Sheet
        2. Click File → Download → CSV
        3. Upload the CSV file here
        """)
        
        uploaded_file = st.file_uploader("Upload CSV file", type=['csv'])
        
        if uploaded_file is not None:
            if st.button("📊 Load Data", type="primary", use_container_width=True):
                with st.spinner("Loading data..."):
                    df, error = load_csv_file(uploaded_file)
                    
                    if df is not None:
                        df = normalize_dataframe(df)
                        st.session_state.data_df = df
                        st.success(f"✅ Loaded {len(df)} records!")
                        st.rerun()
                    else:
                        st.error(f"❌ Error: {error}")
        
        st.divider()
        
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    
    # Main content
    if st.session_state.data_df is not None and not st.session_state.data_df.empty:
        df = st.session_state.data_df
        kpis = calculate_kpis(df)
        
        st.markdown('<div class="section-header"><h2>🏆 Executive Dashboard</h2></div>', unsafe_allow_html=True)
        
        st.markdown(
            render_kpi_row([
                render_kpi("Total Students", f"{kpis['total_students']:,}", "Enrolled", "kpi-box-blue"),
                render_kpi("Started", f"{kpis['started']:,}", f"{kpis['start_rate']}% start rate", "kpi-box-green"),
                render_kpi("Completed", f"{kpis['completed']:,}", f"{kpis['completion_rate']}% completion", "kpi-box-purple"),
                render_kpi("In Progress", f"{kpis['in_progress']:,}", "Active students", "kpi-box-orange"),
            ]),
            unsafe_allow_html=True
        )
        
        st.divider()
        
        # Download Section
        st.markdown("### 📥 Download Options")
        
        col1, col2 = st.columns(2)
        
        with col1:
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📄 Download CSV",
                csv,
                f"service_data_{datetime.now().strftime('%Y%m%d')}.csv",
                "text/csv",
                use_container_width=True
            )
        
        with col2:
            tutor_perf = create_tutor_performance(df)
            team_perf = create_team_performance(df)
            course_analysis = create_course_analysis(df)
            
            excel_data = create_excel_report(df, kpis, tutor_perf, team_perf, course_analysis)
            
            st.download_button(
                "💎 Download Excel Report",
                excel_data,
                f"Service_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        st.divider()
        
        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Overview",
            "👨‍🏫 Tutor Performance",
            "👥 Team Performance",
            "📚 Course Analysis"
        ])
        
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Student Overview</h3></div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Status_Clean' in df.columns:
                    status_counts = df['Status_Clean'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    fig = px.pie(status_counts, values='Count', names='Status', title='Student Status Distribution', hole=0.3)
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                if 'Completion_Status' in df.columns:
                    comp_counts = df['Completion_Status'].value_counts().reset_index()
                    comp_counts.columns = ['Status', 'Count']
                    fig = px.bar(comp_counts, x='Status', y='Count', title='Completion Status', color='Count')
                    st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(df, use_container_width=True, height=300)
        
        with tab2:
            st.markdown('<div class="section-header"><h3>👨‍🏫 Tutor Performance</h3></div>', unsafe_allow_html=True)
            
            if not tutor_perf.empty:
                st.dataframe(tutor_perf, use_container_width=True, height=400)
                fig = px.bar(tutor_perf.head(10), x='Tutor', y='Total_Students',
                           title='Top 10 Tutors by Student Count',
                           color='Completion_Rate_%', color_continuous_scale='Greens')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.markdown('<div class="section-header"><h3>👥 Team Performance</h3></div>', unsafe_allow_html=True)
            
            if not team_perf.empty:
                st.dataframe(team_perf, use_container_width=True, height=400)
                fig = px.bar(team_perf, x='Team', y='Total_Students',
                           title='Team Performance',
                           color='Completion_Rate_%', color_continuous_scale='Blues')
                st.plotly_chart(fig, use_container_width=True)
        
        with tab4:
            st.markdown('<div class="section-header"><h3>📚 Course Analysis</h3></div>', unsafe_allow_html=True)
            
            if not course_analysis.empty:
                st.dataframe(course_analysis, use_container_width=True, height=400)
                fig = px.bar(course_analysis.head(10), x='Course', y='Total_Students',
                           title='Top 10 Courses',
                           color='Completion_Rate_%', color_continuous_scale='Purples')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
    
    else:
        st.info("👈 Upload your CSV file from the sidebar to get started!")
        
        st.markdown("### 📖 Quick Start Guide")
        st.markdown("""
        1. **Export from Google Sheets**: File → Download → CSV
        2. **Upload CSV** using the sidebar
        3. **Click "Load Data"**
        4. **View analytics** across all tabs
        5. **Download reports** as needed
        
        ✅ **No Apps Script deployment needed!**
        """)

if __name__ == "__main__":
    main()
