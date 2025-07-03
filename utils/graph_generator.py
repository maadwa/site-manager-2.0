import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from typing import List, Optional, Dict, Any
import io
import base64

class GraphGenerator:
    """Handles graph generation for various chart types"""
    
    def __init__(self):
        self.color_palette = px.colors.qualitative.Set3
    
    def create_bar_chart(self, df: pd.DataFrame, x_col: str, y_col: str, title: str = None) -> go.Figure:
        """Create a bar chart"""
        try:
            fig = px.bar(
                df, 
                x=x_col, 
                y=y_col,
                title=title or f"{y_col} by {x_col}",
                color_discrete_sequence=self.color_palette
            )
            
            fig.update_layout(
                xaxis_title=x_col,
                yaxis_title=y_col,
                showlegend=False
            )
            
            return fig
        
        except Exception as e:
            print(f"Error creating bar chart: {str(e)}")
            return None
    
    def create_line_chart(self, df: pd.DataFrame, x_col: str, y_col: str, title: str = None) -> go.Figure:
        """Create a line chart"""
        try:
            fig = px.line(
                df, 
                x=x_col, 
                y=y_col,
                title=title or f"{y_col} over {x_col}",
                color_discrete_sequence=self.color_palette
            )
            
            fig.update_layout(
                xaxis_title=x_col,
                yaxis_title=y_col
            )
            
            return fig
        
        except Exception as e:
            print(f"Error creating line chart: {str(e)}")
            return None
    
    def create_scatter_plot(self, df: pd.DataFrame, x_col: str, y_col: str, title: str = None) -> go.Figure:
        """Create a scatter plot"""
        try:
            fig = px.scatter(
                df, 
                x=x_col, 
                y=y_col,
                title=title or f"{y_col} vs {x_col}",
                color_discrete_sequence=self.color_palette
            )
            
            fig.update_layout(
                xaxis_title=x_col,
                yaxis_title=y_col
            )
            
            return fig
        
        except Exception as e:
            print(f"Error creating scatter plot: {str(e)}")
            return None
    
    def create_pie_chart(self, df: pd.DataFrame, values_col: str, names_col: str = None, title: str = None) -> go.Figure:
        """Create a pie chart"""
        try:
            if names_col:
                fig = px.pie(
                    df, 
                    values=values_col, 
                    names=names_col,
                    title=title or f"Distribution of {values_col}",
                    color_discrete_sequence=self.color_palette
                )
            else:
                # Create pie chart from value counts
                value_counts = df[values_col].value_counts()
                fig = px.pie(
                    values=value_counts.values, 
                    names=value_counts.index,
                    title=title or f"Distribution of {values_col}",
                    color_discrete_sequence=self.color_palette
                )
            
            return fig
        
        except Exception as e:
            print(f"Error creating pie chart: {str(e)}")
            return None
    
    def create_histogram(self, df: pd.DataFrame, col: str, bins: int = 20, title: str = None) -> go.Figure:
        """Create a histogram"""
        try:
            fig = px.histogram(
                df, 
                x=col,
                nbins=bins,
                title=title or f"Distribution of {col}",
                color_discrete_sequence=self.color_palette
            )
            
            fig.update_layout(
                xaxis_title=col,
                yaxis_title="Frequency"
            )
            
            return fig
        
        except Exception as e:
            print(f"Error creating histogram: {str(e)}")
            return None
    
    def create_box_plot(self, df: pd.DataFrame, y_col: str, x_col: str = None, title: str = None) -> go.Figure:
        """Create a box plot"""
        try:
            if x_col:
                fig = px.box(
                    df, 
                    x=x_col, 
                    y=y_col,
                    title=title or f"{y_col} by {x_col}",
                    color_discrete_sequence=self.color_palette
                )
            else:
                fig = px.box(
                    df, 
                    y=y_col,
                    title=title or f"Distribution of {y_col}",
                    color_discrete_sequence=self.color_palette
                )
            
            return fig
        
        except Exception as e:
            print(f"Error creating box plot: {str(e)}")
            return None
    
    def create_heatmap(self, df: pd.DataFrame, x_col: str, y_col: str, values_col: str, title: str = None) -> go.Figure:
        """Create a heatmap"""
        try:
            # Create pivot table for heatmap
            pivot_table = df.pivot_table(
                index=y_col, 
                columns=x_col, 
                values=values_col, 
                aggfunc='mean'
            )
            
            fig = px.imshow(
                pivot_table,
                title=title or f"Heatmap: {values_col} by {x_col} and {y_col}",
                color_continuous_scale='Viridis'
            )
            
            return fig
        
        except Exception as e:
            print(f"Error creating heatmap: {str(e)}")
            return None
    
    def create_multi_column_chart(self, df: pd.DataFrame, columns: List[str], chart_type: str = "bar") -> go.Figure:
        """Create chart with multiple columns"""
        try:
            if len(columns) < 2:
                return None
            
            if chart_type == "bar":
                return self.create_bar_chart(df, columns[0], columns[1])
            elif chart_type == "line":
                return self.create_line_chart(df, columns[0], columns[1])
            elif chart_type == "scatter":
                return self.create_scatter_plot(df, columns[0], columns[1])
            elif chart_type == "pie" and len(columns) >= 1:
                return self.create_pie_chart(df, columns[0], columns[1] if len(columns) > 1 else None)
            elif chart_type == "histogram":
                return self.create_histogram(df, columns[0])
            
            return None
        
        except Exception as e:
            print(f"Error creating multi-column chart: {str(e)}")
            return None
    
    def save_figure_as_image(self, fig: go.Figure, filename: str, format: str = "png") -> str:
        """Save figure as image file"""
        try:
            # Save the figure
            fig.write_image(filename, format=format, width=800, height=600, scale=2)
            return filename
        
        except Exception as e:
            print(f"Error saving figure: {str(e)}")
            return None
    
    def figure_to_base64(self, fig: go.Figure, format: str = "png") -> str:
        """Convert figure to base64 string"""
        try:
            img_bytes = fig.to_image(format=format, width=800, height=600, scale=2)
            img_base64 = base64.b64encode(img_bytes).decode()
            return img_base64
        
        except Exception as e:
            print(f"Error converting figure to base64: {str(e)}")
            return None
