"""Streamlit web application for Jira sprint analysis."""

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

from jira_analysis.analyzer import process_discipline_data
from jira_analysis.visualization import create_comparison_plot, format_excel_cell
from jira_analysis.utils import ensure_directory

st.set_page_config(page_title="Jira Sprint Analysis", layout="wide")

st.title("Jira Sprint Analysis Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload your Jira Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Create input directory if it doesn't exist
    ensure_directory("input")
    
    # Save the uploaded file
    input_path = os.path.join("input", uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getvalue())
else:
    # If no file is uploaded, try to use existing file in input directory
    input_files = [f for f in os.listdir("input") if f.endswith('.xlsx')]
    if input_files:
        input_path = os.path.join("input", input_files[0])
    else:
        st.error("No Excel file found. Please upload a file.")
        st.stop()

try:
    # Read the Excel file
    df = pd.read_excel(input_path, skiprows=3)
    
    # Filter out rows with blank or NaN keys
    df = df[df["Key"].notna() & (df["Key"] != "")]
    
    # Display total unique tickets
    st.metric("Total Unique Tickets", len(df["Key"].unique()))
    
    # Display column names for debugging
    st.subheader("Available Columns")
    st.write(list(df.columns))
    
    # Show data preview
    st.subheader("Data Preview")
    st.dataframe(df.head())
    
    # Process button
    if st.button("Process Analysis"):
        with st.spinner("Processing data..."):
            # Ensure output directory exists
            output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
            temp_dir = os.path.join(output_dir, "temp")
            ensure_directory(output_dir)
            ensure_directory(temp_dir)
            
            # Define disciplines
            disciplines = ["QA", "TA", "FE", "BA"]
            
            # Column mapping (same as in original script)
            column_mapping = {
                "QA": {
                    "original": "QA Original Estimate based on old way",
                    "ai": "QA Revised Estimate based on AI",
                    "actual": "QA Actual Time based on AI",
                },
                "TA": {
                    "original": "TA Original Estimate based on old way",
                    "ai": "TA Revised Estimate based on AI",
                    "actual": "TA Actual Time based on AI",
                },
                "FE": {
                    "original": "FE Dev Original Estimate based on old way",
                    "ai": "FE Dev Revised Estimate based on AI",
                    "actual": "FE Dev Actual Time based on AI",
                },
                "BA": {
                    "original": "BA Original Estimate based on old way",
                    "ai": "BA Revised Estimate based on AI",
                    "actual": "BA Actual Time based on AI",
                },
            }
            
            # Calculate total unique tickets
            total_unique_tickets = len(df["Key"].unique())

            # Process data for each discipline
            analysis_results = []
            total_original_estimate = 0
            total_ai_estimate = 0
            total_actual_time = 0
            total_complete = 0
            total_partial = 0
            total_missing = 0

            for discipline, mapping in column_mapping.items():
                discipline_data = process_discipline_data(df, mapping)
                
                # Add to totals
                total_original_estimate += discipline_data["Original Estimate (Total)"]
                total_ai_estimate += discipline_data["AI Estimate (Total)"]
                total_actual_time += discipline_data["Actual Time (Total)"]
                total_complete += discipline_data["Complete Data Points"]
                total_partial += discipline_data["Partial Data Points"]
                total_missing += discipline_data["Missing Data Points"]

                analysis_results.append({
                    "Discipline": discipline,
                    **discipline_data
                })

            # Create metrics display
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Unique Tickets", total_unique_tickets)
                st.metric("Complete Data Points", total_complete)
                st.metric("Partial Data Points", total_partial)
            with col2:
                st.metric("Original Estimate (Hours)", f"{total_original_estimate:.2f}")
                st.metric("AI Estimate (Hours)", f"{total_ai_estimate:.2f}")
                st.metric("Actual Time (Hours)", f"{total_actual_time:.2f}")
            with col3:
                completion_rate = (total_complete / total_unique_tickets * 100) if total_unique_tickets > 0 else 0
                st.metric("Overall Completion Rate", f"{completion_rate:.1f}%")
                st.metric("Missing Data Points", total_missing)
                ai_vs_original = ((total_original_estimate - total_ai_estimate) / total_original_estimate * 100) if total_original_estimate > 0 else 0
                st.metric("AI vs Original Difference", f"{ai_vs_original:.1f}%")

            # Convert results to DataFrame for display
            analysis_df = pd.DataFrame(analysis_results)
            
            # Display summary data in Streamlit
            st.subheader("Summary by Discipline")
            
            # Create a summary table
            summary_data = []
            for discipline in disciplines:
                disc_data = analysis_df[analysis_df['Discipline'] == discipline].iloc[0]
                summary_data.append({
                    'Discipline': discipline,
                    'Complete Data Points': int(disc_data['Complete Data Points']),
                    'Original Estimate': f"{disc_data['Original Estimate (Total)']:.1f}h",
                    'AI Estimate': f"{disc_data['AI Estimate (Total)']:.1f}h",
                    'Actual Time': f"{disc_data['Actual Time (Total)']:.1f}h"
                })
            
            # Display as a DataFrame
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, hide_index=True)
            
            # Create visualization
            try:
                # Estimates Comparison
                fig1 = create_comparison_plot(analysis_df, disciplines)
                
                # Display plot in Streamlit
                st.subheader("Visualization")
                st.pyplot(fig1)
                
                # Save plot for Excel
                st.write("Saving plot...")
                img_bytes1 = BytesIO()
                fig1.savefig(img_bytes1, format='png', bbox_inches="tight", dpi=300)
                img_bytes1.seek(0)
                
                st.write("Adding visualization to Excel...")
                # Visualizations Sheet
                wb = Workbook()
                ws_viz = wb.create_sheet("Visualizations")
                
                try:
                    img1 = Image(img_bytes1)
                    ws_viz.add_image(img1, "A1")
                    st.write("Added estimate comparison plot")
                except Exception as e:
                    st.error(f"Could not add estimate comparison plot: {str(e)}")
                
            except Exception as e:
                st.error(f"Error creating or saving visualization: {str(e)}")
                import traceback
                st.error(traceback.format_exc())
            
            # Summary Sheet (first sheet)
            ws_summary = wb.active
            ws_summary.title = "Summary"
            
            # Define styles
            header_font = Font(bold=True, size=12)
            subheader_font = Font(bold=True)
            normal_font = Font(size=11)
            border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )
            
            # Add title
            ws_summary["A1"] = "Estimation Analysis Summary"
            ws_summary["A1"].font = Font(bold=True, size=14)
            ws_summary.merge_cells("A1:E1")
            ws_summary["A1"].alignment = Alignment(horizontal="center")
            
            # Overall Statistics
            ws_summary["A3"] = "Overall Statistics"
            ws_summary["A3"].font = header_font
            
            # Add overall statistics
            stats_data = [
                ("Total Unique Tickets:", total_unique_tickets),
                ("Total Complete Data Points:", total_complete),
                ("Total Partial Data Points:", total_partial),
                ("Total Missing Data Points:", total_missing),
                ("Overall Completion Rate:", f"{completion_rate:.1f}%"),
                ("Total Original Estimate (Hours):", f"{total_original_estimate:.2f}"),
                ("Total AI Estimate (Hours):", f"{total_ai_estimate:.2f}"),
                ("Total Actual Time (Hours):", f"{total_actual_time:.2f}"),
            ]
            
            for idx, (label, value) in enumerate(stats_data):
                row = idx + 4
                ws_summary[f"A{row}"] = label
                ws_summary[f"B{row}"] = value
                ws_summary[f"A{row}"].font = normal_font
                ws_summary[f"B{row}"].font = normal_font
            
            # Add completion rates by discipline
            ws_summary["A11"] = "Discipline Analysis"
            ws_summary["A11"].font = header_font
            
            headers = [
                "Discipline",
                "Complete",
                "Incomplete",
                "Completion Rate",
                "Original Est. (Total)",
                "AI Est. (Total)",
                "Actual (Total)",
                "AI vs Original Diff %",
            ]
            
            for col, header in enumerate(headers):
                ws_summary.cell(row=12, column=col + 1, value=header).font = subheader_font
            
            for idx, row in analysis_df.iterrows():
                data_row = 13 + idx
                complete = row["Complete Data Points"]
                incomplete = row["Missing Data Points"]
                completion_rate = (
                    (complete / (complete + incomplete) * 100)
                    if (complete + incomplete) > 0
                    else 0
                )
                
                # Calculate AI vs Original difference
                original_est = row["Original Estimate (Total)"]
                ai_est = row["AI Estimate (Total)"]
                ai_vs_original = (
                    ((original_est - ai_est) / original_est * 100)
                    if original_est != 0
                    else 0
                )
                
                ws_summary.cell(row=data_row, column=1, value=row["Discipline"])
                ws_summary.cell(row=data_row, column=2, value=complete)
                ws_summary.cell(row=data_row, column=3, value=incomplete)
                ws_summary.cell(row=data_row, column=4, value=f"{completion_rate:.1f}%")
                ws_summary.cell(row=data_row, column=5, value=original_est).number_format = "0.00"
                ws_summary.cell(row=data_row, column=6, value=ai_est).number_format = "0.00"
                ws_summary.cell(row=data_row, column=7, value=row["Actual Time (Total)"]).number_format = "0.00"
                
                # Add conditional formatting for AI vs Original difference
                diff_cell = ws_summary.cell(row=data_row, column=8, value=f"{ai_vs_original:.1f}%")
                if ai_vs_original > 0:  # AI estimate is lower than original
                    diff_cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green
                else:  # AI estimate is higher than original
                    diff_cell.fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Light red
            
            # Analysis Sheet
            ws_analysis = wb.create_sheet("Analysis")
            
            # Add headers and data to analysis sheet
            for r_idx, r in enumerate(dataframe_to_rows(analysis_df, index=False, header=True)):
                ws_analysis.append(r)
            
            # Original Data Sheet
            ws_original = wb.create_sheet("Original Data")
            
            # Add headers and data to original data sheet
            for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True)):
                ws_original.append(r)
            
            # Generate timestamp for output file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"analysis_results_{timestamp}.xlsx")
            
            # Save Excel file
            wb.save(output_file)
            
            st.success(f"Analysis complete! Results saved to {output_file}")
            
            # Add download button for the generated Excel file
            with open(output_file, "rb") as f:
                st.download_button(
                    label="Download Analysis Results",
                    data=f,
                    file_name=f"analysis_results_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
except Exception as e:
    st.error(f"An error occurred: {str(e)}")
