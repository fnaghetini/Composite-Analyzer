"""
Composite Drillholes Dashboard - A Streamlit application for processing and visualizing drillhole data
using Datamine StudioRM.
"""

import streamlit as st
import helper_functions as hf

# Page configuration
st.set_page_config(
    page_title="Composite Analyzer",
    page_icon="📊",
    layout="wide"
)

# Main function to run the Streamlit app
def main():
    """Main application function."""
    oDmApp = None
    try:
        # Header
        st.title("📊 Composite Analyzer")
        st.markdown("#### Input Parameters")

        # File upload and parameters
        uploaded_file = st.file_uploader(
            "Upload Drillhole DM File",
            type=["dm"],
            help="Select a Datamine format file containing drillhole data"
        )

        if not (uploaded_file):
            st.info("Please upload a file to continue.")
            return
        
        # Display file information
        _, uploaded_dm_name = hf.get_dm_file_path(uploaded_file)
        st.success(f"File {uploaded_dm_name}.dm uploaded successfully!")

        # Initialize Datamine StudioRM
        oDmApp = hf.auto_connect()
        # Get project path
        project_path = hf.get_studio_project_folder_path(oDmApp)

        # Export uploaded DM file to CSV
        hf.export_dm_to_csv(oDmApp, uploaded_dm_name, uploaded_dm_name)
        # Get file columns
        file_columns = hf.get_dm_file_columns(project_path, uploaded_dm_name)

        # Add widget for composite size
        composite_size = st.number_input(
            "Composite Size",
            min_value=1,
            max_value=1000,
            value=1,
            step=1,
            help="Enter the desired composite interval size in meters"
        )

        # Add selectbox for zone and grade fields
        zone_column = st.selectbox("Zone Field", file_columns, help="Select the column representing the zone in the drillhole data")
        grade_column = st.selectbox("Grade Field", file_columns, help="Select the column representing the grade in the drillhole data")

        # Add execute button
        execute_button = st.button(
            "Execute Composite Analysis", 
            help="Click to process the data and generate visualizations",
            type="primary",icon="▶️"
        )

        if not execute_button:
            return

        with st.spinner("Processing data..."):
            # Process composites in Datamine
            output_name = f"{uploaded_dm_name}_composites"
            hf.process_drillhole_composites(oDmApp, uploaded_dm_name, output_name, zone_column, composite_size)

            # Read processed data
            raw_dh = hf.read_csv_file(f"{project_path}/{uploaded_dm_name}.csv")
            composite_dh = hf.read_csv_file(f"{project_path}/{output_name}.csv")      
            
            # Display visualizations with statistics            
            tab1, tab2 = st.tabs(["Length Analysis", "Grade Analysis"])
            
            with tab1:
                st.write("### Length Analysis - Raw vs Composite Samples")
                col1, col2 = st.columns(2)
                
                with col1:
                    fig1 = hf.create_histogram(raw_dh, 'LENGTH', "Raw LENGTH Distribution")
                    st.pyplot(fig1)
                with col2:
                    fig2 = hf.create_histogram(composite_dh, 'LENGTH', f"Composite LENGTH Distribution (size: {composite_size}m)")
                    st.pyplot(fig2)
                
                # Add statistics comparison table for LENGTH
                st.write("### Statistical Comparison - LENGTH")
                length_comparison = hf.create_statistics_comparison(raw_dh, composite_dh, 'LENGTH')
                st.dataframe(length_comparison, use_container_width=True)
            
            with tab2:
                st.write("### Grade Analysis - Raw vs Composite Samples")
                col3, col4 = st.columns(2)
                
                with col3:
                    fig3 = hf.create_histogram(raw_dh, grade_column, f"Raw {grade_column} Distribution", color='goldenrod')
                    st.pyplot(fig3)
                with col4:
                    fig4 = hf.create_histogram(composite_dh, grade_column, f"Composite {grade_column} Distribution (size: {composite_size}m)", color='goldenrod')
                    st.pyplot(fig4)
                
                # Add statistics comparison table for grade
                st.write(f"### Statistical Comparison - {grade_column}")
                grade_comparison = hf.create_statistics_comparison(raw_dh, composite_dh, grade_column)
                st.dataframe(grade_comparison, use_container_width=True)

    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
        st.info("Please check if:\n"
                "1. The file format is correct\n"
                "2. Datamine StudioRM is running\n"
                "3. You have necessary permissions")
    finally:
        if oDmApp is not None:
            hf.cleanup_com_connection()

if __name__ == "__main__":
    main()