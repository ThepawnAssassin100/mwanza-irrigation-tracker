import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from datetime import datetime

# Function to create and export Word document
def to_word(filtered, officer, status, progress):
    doc = Document()
    doc.add_heading('Mwanza District Irrigation Task Report', 0)

    doc.add_paragraph(f"Responsible Officer: {officer}")
    doc.add_paragraph(f"Status: {status}")
    doc.add_paragraph(f"Progress: {progress}")
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    doc.add_heading('Tasks:', level=1)
    table = doc.add_table(rows=1, cols=len(filtered.columns))

    # Adding headers
    for i, col in enumerate(filtered.columns):
        table.cell(0, i).text = col

    # Adding rows
    for _, row in filtered.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()

# Streamlit interface
def main():
    st.set_page_config(page_title="Mwanza Irrigation Tracker", page_icon="🌾", layout="wide")
    
    # Title with Emoji and Styling
    st.markdown(
        """
        <h1 style="color:#4CAF50; font-size:40px; text-align:center;">Mwanza District Irrigation Tracker 🌿</h1>
        <p style="font-size:20px; text-align:center;">Track irrigation tasks and download task reports in Word format.</p>
        """, unsafe_allow_html=True)

    # Data (Replace with your actual data loading process)
    data = {
        'Task': ['Irrigation Setup', 'Field Inspection', 'Maintenance'],
        'Deadline': ['2025-05-01', '2025-05-15', '2025-06-01'],
        'Status': ['In Progress', 'Completed', 'Pending'],
        'Progress': ['50%', '100%', '20%']
    }

    df = pd.DataFrame(data)

    # Show the data in the app with editable fields
    st.write("### Irrigation Tasks List 📅")
    st.write(df)

    # Editable fields for the report
    officer = st.text_input("Responsible Officer", "")
    status = st.selectbox("Status", ["In Progress", "Completed", "Pending"])
    progress = st.text_input("Progress", "")

    # Apply filters for monthly report generation (optional)
    st.write("### Filter Tasks by Month 🔍")
    month_filter = st.date_input("Select the month:", datetime.today())
    filtered_df = df[df['Deadline'].str.contains(month_filter.strftime("%Y-%m"))]

    st.write(f"Filtered tasks for {month_filter.strftime('%B %Y')}:")
    st.write(filtered_df)

    # Add a
