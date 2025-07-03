import streamlit as st
import pandas as pd
import plotly.express as px
from utils.drive_connector import DriveConnector
from utils.excel_processor import ExcelProcessor
import os

def show_project_dashboard():
    """Display project-specific dashboard"""
    st.title("📊 Project Dashboard")
    
    if not st.session_state.get('selected_project'):
        st.warning("Please select a project from the main page")
        return
    
    project = st.session_state.selected_project
    construction_folder = os.getenv("CONSTRUCTION_FOLDER_PATH", "./Construction")
    project_path = os.path.join(construction_folder, project)
    
    st.header(f"Project: {project}")
    
    # Initialize utilities
    drive_connector = DriveConnector()
    excel_processor = ExcelProcessor()
    
    # Get all Excel files in the project
    excel_files = drive_connector.get_excel_files(project_path)
    
    if not excel_files:
        st.warning("No Excel files found in this project")
        return
    
    # Display project overview
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Excel Files", len(excel_files))
    
    total_sheets = 0
    total_records = 0
    
    # Process each file to get statistics
    file_data = []
    
    for file in excel_files:
        file_path = os.path.join(project_path, file)
        sheets = excel_processor.get_sheet_names(file_path)
        sheet_count = len(sheets)
        total_sheets += sheet_count
        
        # Count total records across all sheets
        file_records = 0
        for sheet in sheets:
            try:
                df = excel_processor.load_sheet_data(file_path, sheet)
                if df is not None:
                    file_records += len(df)
            except:
                pass
        
        total_records += file_records
        
        file_data.append({
            'File': file,
            'Sheets': sheet_count,
            'Records': file_records
        })
    
    with col2:
        st.metric("Total Sheets", total_sheets)
    
    with col3:
        st.metric("Total Records", total_records)
    
    # Display file breakdown
    st.subheader("File Breakdown")
    
    if file_data:
        df_files = pd.DataFrame(file_data)
        st.dataframe(df_files, use_container_width=True)
        
        # Visualizations
        col1, col2 = st.columns(2)
        
        with col1:
            fig_sheets = px.bar(
                df_files, 
                x='File', 
                y='Sheets',
                title="Sheets per File"
            )
            fig_sheets.update_xaxes(tickangle=45)
            st.plotly_chart(fig_sheets, use_container_width=True)
        
        with col2:
            fig_records = px.bar(
                df_files, 
                x='File', 
                y='Records',
                title="Records per File"
            )
            fig_records.update_xaxes(tickangle=45)
            st.plotly_chart(fig_records, use_container_width=True)
    
    # File explorer
    st.subheader("File Explorer")
    
    selected_file = st.selectbox(
        "Select file to explore",
        ["Select a file..."] + excel_files
    )
    
    if selected_file != "Select a file...":
        file_path = os.path.join(project_path, selected_file)
        sheets = excel_processor.get_sheet_names(file_path)
        
        if sheets:
            selected_sheet = st.selectbox(
                "Select sheet",
                ["Select a sheet..."] + sheets
            )
            
            if selected_sheet != "Select a sheet...":
                # Display sheet data
                df = excel_processor.load_sheet_data(file_path, selected_sheet)
                
                if df is not None and not df.empty:
                    st.subheader(f"Sheet: {selected_sheet}")
                    
                    # Sheet statistics
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Rows", len(df))
                    with col2:
                        st.metric("Columns", len(df.columns))
                    with col3:
                        st.metric("Numeric Columns", len(df.select_dtypes(include=['number']).columns))
                    with col4:
                        st.metric("Non-null Values", df.count().sum())
                    
                    # Display data sample
                    st.subheader("Data Preview")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    # Column information
                    st.subheader("Column Information")
                    
                    column_info = []
                    for col in df.columns:
                        column_info.append({
                            'Column': col,
                            'Type': str(df[col].dtype),
                            'Non-null': df[col].count(),
                            'Unique': df[col].nunique(),
                            'Sample': str(df[col].iloc[0]) if len(df) > 0 else 'N/A'
                        })
                    
                    df_columns = pd.DataFrame(column_info)
                    st.dataframe(df_columns, use_container_width=True)
                    
                    # Quick visualization
                    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
                    if len(numeric_cols) >= 1:
                        st.subheader("Quick Visualization")
                        
                        if len(numeric_cols) >= 2:
                            col1, col2 = st.columns(2)
                            with col1:
                                x_col = st.selectbox("X-axis", numeric_cols, key="x_axis")
                            with col2:
                                y_col = st.selectbox("Y-axis", numeric_cols, index=1 if len(numeric_cols) > 1 else 0, key="y_axis")
                            
                            if x_col and y_col:
                                fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
                                st.plotly_chart(fig, use_container_width=True)
                        
                        # Single column histogram
                        selected_col = st.selectbox("Select column for histogram", numeric_cols, key="hist_col")
                        if selected_col:
                            fig_hist = px.histogram(df, x=selected_col, title=f"Distribution of {selected_col}")
                            st.plotly_chart(fig_hist, use_container_width=True)

if __name__ == "__main__":
    show_project_dashboard()
