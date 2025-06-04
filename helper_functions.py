"""Helper functions for Datamine StudioRM Drillholes processing."""

import pythoncom
import win32com.client as client
import pandas as pd
import os
import streamlit as st
from typing import Tuple
import matplotlib.pyplot as plt


def auto_connect():
    """Initialize connection with Datamine StudioRM or return None in demo mode."""
    try:
        # Check if we're in demo mode (e.g., when deployed)
        if os.environ.get('DEMO_MODE', 'false').lower() == 'true':
            st.warning("Running in demo mode - Datamine StudioRM features are limited")
            return None
            
        pythoncom.CoInitialize()
        return client.Dispatch("Datamine.StudioRM.Application")
    except Exception as e:
        st.error("Failed to connect to Datamine StudioRM. Running in demo mode.")
        return None

def get_dm_file_path(uploaded_file) -> Tuple[str, str]:
    """
    Process uploaded DM file and return its path and name.
    
    Args:
        uploaded_file: Streamlit uploaded file object
    Returns:
        Tuple containing file path and name
    """
    file_name = uploaded_file.name[:-3]
    os.makedirs("uploaded_files", exist_ok=True)
    file_path = os.path.join("uploaded_files", file_name) + ".dm"
    
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return file_path, file_name


def get_studio_project_folder_path(oDmApp) -> str:
    """Get the current Datamine StudioRM project folder path."""
    try:
        return oDmApp.ActiveProject.Folder.replace("\\", "/")
    except Exception as e:
        st.error("Failed to get project folder path from StudioRM")
        raise e


def read_csv_file(file_path: str) -> pd.DataFrame:
    """
    Read CSV file into pandas DataFrame with error handling.
    
    Args:
        file_path: Path to the CSV file
    Returns:
        pandas DataFrame with the file contents
    """
    try:
        df = pd.read_csv(file_path)
        if df.empty:
            raise ValueError("Empty file")
        return df
    except Exception as e:
        st.error(f"Error reading CSV file: {str(e)}")
        raise e


def create_histogram(data: pd.DataFrame, column: str, title: str, color: str = 'lightgray') -> plt.Figure:
    """
    Create a histogram plot with descriptive statistics for the given data.
    
    Args:
        data: DataFrame containing the data
        column: Column name to plot
        title: Title for the plot
        color: Color for the histogram bars
    Returns:
        matplotlib Figure object
    """
    # Calculate statistics
    stats = [
        ('# Samples', f"{len(data[column]):,}"),
        ('Sum', f"{data[column].sum():,.4f}"),
        ('Mean', f"{data[column].mean():,.4f}"),
        ('Median', f"{data[column].median():,.4f}"),
        ('Variance', f"{data[column].var():,.4f}"),
        ('Std. Dev.', f"{data[column].std():,.4f}")
    ]
    
    # Create figure
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Create histogram
    ax.hist(data[column], bins=30, color=color, edgecolor='black')
    ax.set_title(title, pad=20)
    ax.set_xlabel(column)
    ax.set_ylabel('Frequency')
      # Add statistics text box
    stats_text = '\n'.join([f"{name}: {value}" for name, value in stats])
    ax.text(0.95, 0.95, stats_text,
            transform=ax.transAxes,
            verticalalignment='top',
            horizontalalignment='right',
            bbox=dict(boxstyle='round,pad=0.8',
                     facecolor='white',
                     edgecolor='gray',
                     alpha=0.9),
            fontsize=10,
            family='monospace')
    
    plt.tight_layout()
    return fig

def export_dm_to_csv(oDmApp, input_file: str, output_file: str):
    """
    Export Datamine file to CSV format.
    
    Args:
        oDmApp: Datamine StudioRM application object
        input_file: Input file name
        output_file: Output file name
    """
    export_command = (
        f"output &IN={input_file} @CSV=1 @NODD=0 "
        f"@DPLACE=-1 @IMPLICIT=1 '{output_file}.csv'"
    )
    
    try:
        oDmApp.ParseCommand(export_command)
    except Exception as e:
        st.error(f"Error exporting DM file {input_file} to CSV")
        raise e
    
def get_dm_file_columns(project_path: str, file_name: str) -> list:
    """
    Get columns of the uploaded DM file.
    
    Args:
        project_path: Path to the Datamine project folder
        uploaded_file: Streamlit uploaded file object
    Returns:
        List of column names in the DM file
    """
    try:
        file_path = f"{project_path}/{file_name}.csv"
        file_columns = read_csv_file(file_path).columns.tolist()
        return file_columns
    except Exception as e:
        st.error(f"Error getting columns from DM file: {str(e)}")
        raise e
    

def process_drillhole_composites(oDmApp, input_file: str, output_file: str, zone_col: str, composite_size: float):
    """
    Process drillhole composites using Datamine StudioRM.
    
    Args:
        input_file: Input file name
        output_file: Output file name
        zone_col: Column representing the zone in the drillhole data
        composite_size: Size for compositing
    """
    
    composite_command = (
        f"compdh &IN={input_file} &OUT={output_file} "
        f"*BHID=BHID *FROM=FROM *TO=TO *ZONE={zone_col} "
        f"@INTERVAL={str(composite_size)} @MAXGAP=0 @START=0 @MODE=1 @PRINT=0"
    )
        
    try:
        oDmApp.ParseCommand(composite_command)
        export_dm_to_csv(oDmApp, output_file, output_file)
    except Exception as e:
        st.error("Error processing drillhole composites in StudioRM")
        raise e


def cleanup_com_connection():
    """
    Clean up COM connection with Datamine StudioRM.
    Should be called at the end of the application.
    """
    try:
        pythoncom.CoUninitialize()
    except Exception as e:
        st.error(f"Error closing COM connection: {str(e)}")

def create_statistics_comparison(raw_data: pd.DataFrame, composite_data: pd.DataFrame, column: str) -> pd.DataFrame:
    """
    Create a comparison table between raw and composite statistics.
    
    Args:
        raw_data: DataFrame containing raw samples
        composite_data: DataFrame containing composite samples
        column: Column name to analyze
    Returns:
        DataFrame with statistics comparison
    """
    # Calculate statistics for both datasets
    raw_stats = {
        'Nb. of Samples': len(raw_data[column]),
        'Sum': raw_data[column].sum(),
        'Mean': raw_data[column].mean(),
        'Median': raw_data[column].median(),
        'Variance': raw_data[column].var(),
        'Std Dev': raw_data[column].std()
    }
    
    comp_stats = {
        'Nb. of Samples': len(composite_data[column]),
        'Sum': composite_data[column].sum(),
        'Mean': composite_data[column].mean(),
        'Median': composite_data[column].median(),
        'Variance': composite_data[column].var(),
        'Std Dev': composite_data[column].std()
    }
    
    # Calculate percentage differences
    diff_stats = {}
    for key in raw_stats:
        if raw_stats[key] != 0:  # Avoid division by zero
            diff_pct = round(((comp_stats[key] - raw_stats[key]) / raw_stats[key]) * 100, 2)
        else:
            diff_pct = 0
        diff_stats[key] = diff_pct
    
    # Create comparison DataFrame
    comparison_df = pd.DataFrame({
        'Raw': raw_stats,
        'Composite': comp_stats,
        'Difference (%)': diff_stats
    }).round(4)
    
    return comparison_df