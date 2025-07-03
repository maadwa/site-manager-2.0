import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from utils.excel_processor import ExcelProcessor
from utils.graph_generator import GraphGenerator
from utils.ppt_generator import PPTGenerator
import os

def show_statistics_viewer():
    """Display statistics viewer page"""
    st.title("📈 Statistics Viewer")
    
    # Check if we have required session state
    if not all([
        st.session_state.get('selected_project'),
        st.session_state.get('selected_file'),
        st.session_state.get('selected_sheet')
    ]):
        st.warning("Please select a project, file, and sheet from the main page")
        return
    
    project = st.session_state.selected_project
    file = st.session_state.selected_file
    sheet = st.session_state.selected_sheet
    
    construction_folder = os.getenv("CONSTRUCTION_FOLDER_PATH", "./Construction")
    file_path = os.path.join(construction_folder, project, file)
    
    st.header(f"Statistics for: {project} / {file} / {sheet}")
    
    # Initialize utilities
    excel_processor = ExcelProcessor()
    graph_generator = GraphGenerator()
    
    try:
        # Load data
        df = excel_processor.load_sheet_data(file_path, sheet)
        
        if df is None or df.empty:
            st.error("No data found in the selected sheet")
            return
        
        # Display basic statistics
        st.subheader("📊 Basic Statistics")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df))
        with col2:
            st.metric("Total Columns", len(df.columns))
        with col3:
            st.metric("Numeric Columns", len(df.select_dtypes(include=['number']).columns))
        with col4:
            st.metric("Missing Values", df.isnull().sum().sum())
        
        # Column selection for analysis
        st.subheader("🎯 Column Selection for Analysis")
        
        # Get different types of columns
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        datetime_cols = df.select_dtypes(include=['datetime']).columns.tolist()
        all_cols = df.columns.tolist()
        
        # Multi-select for columns
        selected_columns = st.multiselect(
            "Select columns for analysis",
            all_cols,
            default=numeric_cols[:2] if len(numeric_cols) >= 2 else numeric_cols,
            help="Choose columns you want to analyze and visualize"
        )
        
        if selected_columns:
            # Display selected columns info
            st.subheader("📋 Selected Columns Information")
            
            col_info = []
            for col in selected_columns:
                col_info.append({
                    'Column': col,
                    'Type': str(df[col].dtype),
                    'Non-null Count': df[col].count(),
                    'Unique Values': df[col].nunique(),
                    'Min': df[col].min() if pd.api.types.is_numeric_dtype(df[col]) else 'N/A',
                    'Max': df[col].max() if pd.api.types.is_numeric_dtype(df[col]) else 'N/A',
                    'Mean': round(df[col].mean(), 2) if pd.api.types.is_numeric_dtype(df[col]) else 'N/A'
                })
            
            df_col_info = pd.DataFrame(col_info)
            st.dataframe(df_col_info, use_container_width=True)
            
            # Visualization options
            st.subheader("📈 Create Visualizations")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                chart_type = st.selectbox(
                    "Chart Type",
                    ["Bar Chart", "Line Chart", "Scatter Plot", "Pie Chart", "Histogram", "Box Plot"]
                )
            
            with col2:
                if len(selected_columns) >= 2:
                    x_axis = st.selectbox("X-axis", selected_columns, key="x_axis_stats")
                    y_axis = st.selectbox("Y-axis", selected_columns, index=1, key="y_axis_stats")
                else:
                    x_axis = selected_columns[0] if selected_columns else None
                    y_axis = None
                    st.info("Select at least 2 columns for X and Y axis selection")
            
            with col3:
                if st.button("🎨 Generate Visualization"):
                    create_visualization(df, selected_columns, chart_type, x_axis, y_axis, graph_generator)
            
            # Pre-generated visualizations for all selected columns
            if len(selected_columns) >= 1:
                st.subheader("🎯 Automatic Visualizations")
                
                # Create tabs for different visualization types
                tab1, tab2, tab3, tab4 = st.tabs(["Distribution", "Relationships", "Summary", "Trends"])
                
                with tab1:
                    create_distribution_charts(df, selected_columns)
                
                with tab2:
                    create_relationship_charts(df, selected_columns)
                
                with tab3:
                    create_summary_charts(df, selected_columns)
                
                with tab4:
                    create_trend_charts(df, selected_columns)
            
            # Statistics sheet management
            st.subheader("📊 Statistics Sheet Management")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📝 Create/Update Statistics Sheet"):
                    with st.spinner("Creating statistics sheet..."):
                        success = excel_processor.create_statistics_sheet(file_path, df, selected_columns)
                        if success:
                            st.success("Statistics sheet created/updated successfully!")
                        else:
                            st.error("Failed to create statistics sheet")
            
            with col2:
                # Check if statistics sheet exists
                has_stats = excel_processor.has_statistics_sheet(file_path)
                if has_stats:
                    st.success("✅ Statistics sheet exists")
                else:
                    st.info("ℹ️ No statistics sheet found")
            
            # PowerPoint export
            st.subheader("📋 Export to PowerPoint")
            
            export_options = st.multiselect(
                "Select visualizations to include in PowerPoint",
                ["Distribution Charts", "Relationship Charts", "Summary Charts", "Trend Charts"],
                default=["Summary Charts"]
            )
            
            if st.button("📋 Generate PowerPoint Report"):
                if export_options:
                    generate_powerpoint_report(df, selected_columns, export_options, sheet, graph_generator)
                else:
                    st.warning("Please select at least one visualization type to include")
        
        else:
            st.info("Please select columns to begin analysis")
        
        # Raw data viewer
        with st.expander("👁️ View Raw Data"):
            st.dataframe(df, use_container_width=True)
            
            # Download options
            col1, col2 = st.columns(2)
            with col1:
                csv = df.to_csv(index=False)
                st.download_button(
                    label="📄 Download as CSV",
                    data=csv,
                    file_name=f"{project}_{file}_{sheet}.csv",
                    mime="text/csv"
                )
            
            with col2:
                # Note: Excel download would require additional processing
                st.info("Excel download available through file system")
    
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")

def create_visualization(df, selected_columns, chart_type, x_axis, y_axis, graph_generator):
    """Create and display visualization based on user selection"""
    try:
        fig = None
        
        if chart_type == "Bar Chart" and x_axis and y_axis:
            fig = graph_generator.create_bar_chart(df, x_axis, y_axis)
        elif chart_type == "Line Chart" and x_axis and y_axis:
            fig = graph_generator.create_line_chart(df, x_axis, y_axis)
        elif chart_type == "Scatter Plot" and x_axis and y_axis:
            fig = graph_generator.create_scatter_plot(df, x_axis, y_axis)
        elif chart_type == "Pie Chart" and x_axis:
            fig = graph_generator.create_pie_chart(df, x_axis, y_axis)
        elif chart_type == "Histogram" and x_axis:
            fig = graph_generator.create_histogram(df, x_axis)
        elif chart_type == "Box Plot" and x_axis:
            fig = graph_generator.create_box_plot(df, x_axis, y_axis)
        
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Unable to create visualization with selected parameters")
    
    except Exception as e:
        st.error(f"Error creating visualization: {str(e)}")

def create_distribution_charts(df, selected_columns):
    """Create distribution charts for selected columns"""
    numeric_cols = [col for col in selected_columns if pd.api.types.is_numeric_dtype(df[col])]
    
    if numeric_cols:
        for col in numeric_cols:
            fig = px.histogram(df, x=col, title=f"Distribution of {col}")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric columns selected for distribution analysis")

def create_relationship_charts(df, selected_columns):
    """Create relationship charts between selected columns"""
    numeric_cols = [col for col in selected_columns if pd.api.types.is_numeric_dtype(df[col])]
    
    if len(numeric_cols) >= 2:
        # Correlation heatmap
        if len(numeric_cols) > 2:
            corr_matrix = df[numeric_cols].corr()
            fig = px.imshow(corr_matrix, title="Correlation Matrix")
            st.plotly_chart(fig, use_container_width=True)
        
        # Scatter plots for pairs
        for i in range(len(numeric_cols)):
            for j in range(i+1, len(numeric_cols)):
                fig = px.scatter(df, x=numeric_cols[i], y=numeric_cols[j], 
                               title=f"{numeric_cols[j]} vs {numeric_cols[i]}")
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Need at least 2 numeric columns for relationship analysis")

def create_summary_charts(df, selected_columns):
    """Create summary charts for selected columns"""
    numeric_cols = [col for col in selected_columns if pd.api.types.is_numeric_dtype(df[col])]
    
    if numeric_cols:
        # Summary statistics
        summary_stats = df[numeric_cols].describe()
        fig = go.Figure()
        
        for col in numeric_cols:
            fig.add_trace(go.Bar(
                x=['Mean', 'Std', 'Min', 'Max'],
                y=[summary_stats.loc['mean', col], summary_stats.loc['std', col], 
                   summary_stats.loc['min', col], summary_stats.loc['max', col]],
                name=col
            ))
        
        fig.update_layout(title="Summary Statistics Comparison", barmode='group')
        st.plotly_chart(fig, use_container_width=True)
        
        # Box plots
        fig_box = go.Figure()
        for col in numeric_cols:
            fig_box.add_trace(go.Box(y=df[col], name=col))
        fig_box.update_layout(title="Box Plot Comparison")
        st.plotly_chart(fig_box, use_container_width=True)

def create_trend_charts(df, selected_columns):
    """Create trend charts if there are date columns"""
    date_cols = [col for col in df.columns if 'date' in col.lower() or pd.api.types.is_datetime64_any_dtype(df[col])]
    numeric_cols = [col for col in selected_columns if pd.api.types.is_numeric_dtype(df[col])]
    
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        for num_col in numeric_cols:
            try:
                fig = px.line(df, x=date_col, y=num_col, title=f"{num_col} over time")
                st.plotly_chart(fig, use_container_width=True)
            except:
                pass
    else:
        st.info("No date columns found for trend analysis")

def generate_powerpoint_report(df, selected_columns, export_options, sheet_name, graph_generator):
    """Generate PowerPoint report with selected visualizations"""
    try:
        with st.spinner("Generating PowerPoint report..."):
            ppt_generator = PPTGenerator()
            
            graphs = []
            
            # Generate graphs based on export options
            numeric_cols = [col for col in selected_columns if pd.api.types.is_numeric_dtype(df[col])]
            
            if "Distribution Charts" in export_options and numeric_cols:
                for col in numeric_cols[:3]:  # Limit to first 3 to avoid too many slides
                    fig = px.histogram(df, x=col, title=f"Distribution of {col}")
                    graphs.append({
                        'title': f"Distribution of {col}",
                        'figure': fig
                    })
            
            if "Relationship Charts" in export_options and len(numeric_cols) >= 2:
                fig = px.scatter(df, x=numeric_cols[0], y=numeric_cols[1], 
                               title=f"{numeric_cols[1]} vs {numeric_cols[0]}")
                graphs.append({
                    'title': f"Relationship: {numeric_cols[1]} vs {numeric_cols[0]}",
                    'figure': fig
                })
            
            if "Summary Charts" in export_options and numeric_cols:
                # Box plot
                fig_box = go.Figure()
                for col in numeric_cols:
                    fig_box.add_trace(go.Box(y=df[col], name=col))
                fig_box.update_layout(title="Summary Statistics")
                graphs.append({
                    'title': "Summary Statistics - Box Plot",
                    'figure': fig_box
                })
            
            if graphs:
                ppt_path = ppt_generator.create_presentation(graphs, f"{sheet_name}_Statistics_Report")
                
                if ppt_path and os.path.exists(ppt_path):
                    st.success("PowerPoint report generated successfully!")
                    
                    # Provide download
                    with open(ppt_path, "rb") as file:
                        st.download_button(
                            label="📥 Download PowerPoint Report",
                            data=file.read(),
                            file_name=f"{sheet_name}_Statistics_Report.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                else:
                    st.error("Failed to generate PowerPoint report")
            else:
                st.warning("No graphs generated for the selected options")
    
    except Exception as e:
        st.error(f"Error generating PowerPoint report: {str(e)}")

if __name__ == "__main__":
    show_statistics_viewer()
