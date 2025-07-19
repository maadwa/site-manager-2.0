# Construction Project Management Dashboard

## Overview

This is a Streamlit-based web application for managing and analyzing construction project data. The application provides a dashboard interface for viewing Excel files stored in construction project folders, generating statistics, visualizations, and PowerPoint reports.

## System Architecture

### Frontend Architecture
- **Framework**: Streamlit for web interface
- **Layout**: Multi-page application with sidebar navigation
- **State Management**: Streamlit session state for maintaining user selections across pages
- **Visualization**: Plotly for interactive charts and graphs

### Backend Architecture
- **Structure**: Modular utility classes for different functionalities
- **File Processing**: Direct file system access for Excel file operations
- **Data Processing**: Pandas for data manipulation and analysis
- **Report Generation**: Python-pptx for PowerPoint presentation creation

### Data Storage Solutions
- **Primary Storage**: File system based (shared drive or local folders)
- **File Format**: Excel files (.xlsx, .xls, .xlsm) as primary data source
- **Structure**: Hierarchical folder structure (Construction > Projects > Excel files)

## Key Components

### Main Application (`app.py`)
- Entry point and main dashboard
- Project selection interface
- Drive connectivity and initial setup
- Session state initialization

### Page Components
- **Project Dashboard** (`pages/project_dashboard.py`): Project-level overview and metrics
- **Statistics Viewer** (`pages/statistics_viewer.py`): Detailed data analysis and visualization

### Utility Classes
- **DriveConnector** (`utils/drive_connector.py`): File system operations and folder navigation
- **ExcelProcessor** (`utils/excel_processor.py`): Excel file reading, sheet processing, and data cleaning
- **GraphGenerator** (`utils/graph_generator.py`): Chart and visualization creation using Plotly
- **PPTGenerator** (`utils/ppt_generator.py`): PowerPoint presentation generation

## Data Flow

1. **Project Discovery**: Application scans construction folder for project directories
2. **File Detection**: Identifies Excel files within selected project folders
3. **Sheet Processing**: Extracts sheet names and loads data from selected sheets
4. **Data Cleaning**: Processes and cleans data (removes empty rows/columns, converts dates)
5. **Analysis**: Generates statistics and creates visualizations
6. **Report Generation**: Creates downloadable PowerPoint presentations with charts

## External Dependencies

### Core Libraries
- **streamlit**: Web application framework
- **pandas**: Data manipulation and analysis
- **plotly**: Interactive visualization library
- **openpyxl**: Excel file processing
- **python-pptx**: PowerPoint generation

### File System Requirements
- Shared drive access or local file system
- Environment variable `CONSTRUCTION_FOLDER_PATH` for custom folder location
- Default folder structure: `./Construction/[Project Name]/[Excel Files]`

## Deployment Strategy

### Environment Setup
- Python environment with required dependencies
- File system access to construction project folders
- Optional environment variable configuration for custom paths

### Configuration
- Environment variable `CONSTRUCTION_FOLDER_PATH` for specifying construction folder location
- Default fallback to `./Construction` directory
- Wide layout configuration for optimal dashboard viewing

### Error Handling
- Graceful handling of missing folders or files
- User-friendly error messages and warnings
- Fallback mechanisms for file access issues

## Changelog
- July 03, 2025. Initial setup

## User Preferences

Preferred communication style: Simple, everyday language.
