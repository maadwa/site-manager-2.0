import streamlit as st
import os
import pandas as pd
from utils.drive_connector import DriveConnector
from utils.excel_processor import ExcelProcessor
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Construction Project Management",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'selected_project' not in st.session_state:
    st.session_state.selected_project = None
if 'selected_file' not in st.session_state:
    st.session_state.selected_file = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None

def main():
    st.title("🏗️ Construction Project Management Dashboard")
    
    # Initialize drive connector
    drive_connector = DriveConnector()
    excel_processor = ExcelProcessor()
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    
    # Get construction folder path from environment or use default
    construction_folder = os.getenv("CONSTRUCTION_FOLDER_PATH", "./Construction")
    
    if not os.path.exists(construction_folder):
        st.error(f"Construction folder not found at: {construction_folder}")
        st.info("Please ensure the shared drive is mounted or set CONSTRUCTION_FOLDER_PATH environment variable")
        return
    
    # Get project folders
    try:
        projects = drive_connector.get_project_folders(construction_folder)
        
        if not projects:
            st.warning("No project folders found in the Construction directory")
            return
            
        # Project selection
        selected_project = st.sidebar.selectbox(
            "Select Project",
            ["Select a project..."] + projects,
            key="project_selector"
        )
        
        if selected_project != "Select a project...":
            st.session_state.selected_project = selected_project
            project_path = os.path.join(construction_folder, selected_project)
            
            # Get Excel files in the project
            excel_files = drive_connector.get_excel_files(project_path)
            
            if excel_files:
                selected_file = st.sidebar.selectbox(
                    "Select Excel File",
                    ["Select a file..."] + excel_files,
                    key="file_selector"
                )
                
                if selected_file != "Select a file...":
                    st.session_state.selected_file = selected_file
                    file_path = os.path.join(project_path, selected_file)
                    
                    # Get sheets in the Excel file
                    sheets = excel_processor.get_sheet_names(file_path)
                    
                    if sheets:
                        selected_sheet = st.sidebar.selectbox(
                            "Select Sheet",
                            ["Select a sheet..."] + sheets,
                            key="sheet_selector"
                        )
                        
                        if selected_sheet != "Select a sheet...":
                            st.session_state.selected_sheet = selected_sheet
                            display_sheet_analysis(file_path, selected_sheet, excel_processor)
            else:
                st.sidebar.warning("No Excel files found in this project")
        
        # Display overall summary
        display_construction_summary(construction_folder, drive_connector)
        
    except Exception as e:
        st.error(f"Error accessing construction folder: {str(e)}")

def display_construction_summary(construction_folder, drive_connector):
    """Display overall construction project summary"""
    st.header("📊 Construction Folder Summary")
    
    try:
        projects = drive_connector.get_project_folders(construction_folder)
        
        if projects:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Projects", len(projects))
            
            total_files = 0
            total_sheets = 0
            
            project_data = []
            
            for project in projects:
                project_path = os.path.join(construction_folder, project)
                excel_files = drive_connector.get_excel_files(project_path)
                file_count = len(excel_files)
                total_files += file_count
                
                sheet_count = 0
                for file in excel_files:
                    try:
                        file_path = os.path.join(project_path, file)
                        excel_processor = ExcelProcessor()
                        sheets = excel_processor.get_sheet_names(file_path)
                        sheet_count += len(sheets)
                    except:
                        pass
                
                total_sheets += sheet_count
                project_data.append({
                    'Project': project,
                    'Excel Files': file_count,
                    'Total Sheets': sheet_count
                })
            
            with col2:
                st.metric("Total Excel Files", total_files)
            
            with col3:
                st.metric("Total Sheets", total_sheets)
            
            # Project breakdown table
            if project_data:
                st.subheader("Project Breakdown")
                df = pd.DataFrame(project_data)
                st.dataframe(df, use_container_width=True)
                
                # Visual charts
                col1, col2 = st.columns(2)
                
                with col1:
                    fig_files = px.bar(
                        df, 
                        x='Project', 
                        y='Excel Files',
                        title="Excel Files per Project"
                    )
                    st.plotly_chart(fig_files, use_container_width=True)
                
                with col2:
                    fig_sheets = px.pie(
                        df, 
                        values='Total Sheets', 
                        names='Project',
                        title="Sheet Distribution by Project"
                    )
                    st.plotly_chart(fig_sheets, use_container_width=True)
        
    except Exception as e:
        st.error(f"Error generating summary: {str(e)}")

def display_sheet_analysis(file_path, sheet_name, excel_processor):
    """Display analysis for selected sheet"""
    st.header(f"📈 Analysis: {sheet_name}")
    
    try:
        # Load data
        df = excel_processor.load_sheet_data(file_path, sheet_name)
        
        if df is not None and not df.empty:
            # Display basic info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Rows", len(df))
            with col2:
                st.metric("Columns", len(df.columns))
            with col3:
                st.metric("Non-null Values", df.count().sum())
            
            # Column selection for visualization
            st.subheader("📊 Create Visualizations")
            
            # Get numeric and categorical columns
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
            all_cols = df.columns.tolist()
            
            col1, col2 = st.columns(2)
            
            with col1:
                selected_columns = st.multiselect(
                    "Select columns for visualization",
                    all_cols,
                    help="Choose columns to create graphs"
                )
            
            with col2:
                chart_type = st.selectbox(
                    "Chart Type",
                    ["Bar Chart", "Line Chart", "Scatter Plot", "Pie Chart", "Histogram"]
                )
            
            # Generate visualization
            if selected_columns:
                generate_visualization(df, selected_columns, chart_type)
            
            # Statistics creation button
            if st.button("📊 Create/Update Statistics Sheet"):
                try:
                    success = excel_processor.create_statistics_sheet(file_path, df, selected_columns)
                    if success:
                        st.success("Statistics sheet created/updated successfully!")
                        st.rerun()
                    else:
                        st.error("Failed to create statistics sheet")
                except Exception as e:
                    st.error(f"Error creating statistics sheet: {str(e)}")
            
            # PowerPoint export
            if st.button("📋 Generate PowerPoint Report"):
                try:
                    from utils.ppt_generator import PPTGenerator
                    ppt_gen = PPTGenerator()
                    
                    # Generate graphs for PowerPoint
                    graphs = []
                    if selected_columns and len(selected_columns) >= 2:
                        fig = create_chart(df, selected_columns, chart_type)
                        graphs.append({
                            'title': f"{chart_type} - {', '.join(selected_columns)}",
                            'figure': fig
                        })
                    
                    if graphs:
                        ppt_path = ppt_gen.create_presentation(graphs, f"{sheet_name}_Report")
                        st.success(f"PowerPoint report generated: {ppt_path}")
                        
                        # Provide download link
                        with open(ppt_path, "rb") as file:
                            st.download_button(
                                label="Download PowerPoint",
                                data=file.read(),
                                file_name=f"{sheet_name}_Report.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                    else:
                        st.warning("Please select at least 2 columns to generate PowerPoint report")
                        
                except Exception as e:
                    st.error(f"Error generating PowerPoint: {str(e)}")
            
            # Display raw data
            with st.expander("View Raw Data"):
                st.dataframe(df, use_container_width=True)
                
        else:
            st.warning("No data found in the selected sheet")
            
    except Exception as e:
        st.error(f"Error loading sheet data: {str(e)}")

def generate_visualization(df, selected_columns, chart_type):
    """Generate visualization based on selected columns and chart type"""
    try:
        if len(selected_columns) < 1:
            st.warning("Please select at least one column")
            return
        
        fig = create_chart(df, selected_columns, chart_type)
        
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error creating visualization: {str(e)}")

def create_chart(df, selected_columns, chart_type):
    """Create chart based on type and columns"""
    try:
        if chart_type == "Bar Chart":
            if len(selected_columns) >= 2:
                return px.bar(df, x=selected_columns[0], y=selected_columns[1])
            else:
                return px.bar(df, y=selected_columns[0])
                
        elif chart_type == "Line Chart":
            if len(selected_columns) >= 2:
                return px.line(df, x=selected_columns[0], y=selected_columns[1])
            else:
                return px.line(df, y=selected_columns[0])
                
        elif chart_type == "Scatter Plot":
            if len(selected_columns) >= 2:
                return px.scatter(df, x=selected_columns[0], y=selected_columns[1])
            else:
                st.warning("Scatter plot requires at least 2 columns")
                return None
                
        elif chart_type == "Pie Chart":
            if len(selected_columns) >= 1:
                # Use value counts for categorical data
                if df[selected_columns[0]].dtype == 'object':
                    value_counts = df[selected_columns[0]].value_counts()
                    return px.pie(values=value_counts.values, names=value_counts.index)
                else:
                    return px.pie(df, values=selected_columns[0])
            
        elif chart_type == "Histogram":
            if len(selected_columns) >= 1:
                return px.histogram(df, x=selected_columns[0])
                
        return None
        
    except Exception as e:
        st.error(f"Error creating {chart_type}: {str(e)}")
        return None

if __name__ == "__main__":
    main()
