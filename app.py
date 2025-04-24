import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from datetime import datetime
from PIL import Image
import io

# Function to create and export Excel file
def to_excel(filtered, officer, location, epa, ta, image_data=None):
    output = BytesIO()
    
    # Add the extra information to the DataFrame
    filtered['Responsible Officer'] = officer
    filtered['Location'] = location
    filtered['EPA'] = epa
    filtered['T/A'] = ta

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Sheet1')

    output.seek(0)
    return output.getvalue()

# Function to create a Word document
def to_word(filtered, officer, location, epa, ta, image_data=None):
    doc = Document()
    doc.add_heading('Mwanza District Irrigation Task Report', 0)

    doc.add_paragraph(f"Responsible Officer: {officer}")
    doc.add_paragraph(f"Location: {location}")
    doc.add_paragraph(f"EPA: {epa}")
    doc.add_paragraph(f"T/A: {ta}")
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    if image_data:
        doc.add_paragraph("Image for Monthly Report:")
        doc.add_picture(image_data)

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
        <p style="font-size:20px; text-align:center;">Track irrigation tasks and download task reports in Excel and Word formats.</p>
        """, unsafe_allow_html=True)

    # Data (Replace with your actual data loading process)
    data = {
        'Task': ['Irrigation Setup', 'Field Inspection', 'Maintenance'],
        'Deadline': ['2025-05-01', '2025-05-15', '2025-06-01'],
        'Status': ['In Progress', 'Completed', 'Pending']
    }

    df = pd.DataFrame(data)

    # Show the data in the app with editable fields
    st.write("### Irrigation Tasks List 📅")
    st.write(df)

    # Editable fields for the report
    officer = st.text_input("Responsible Officer", "")
    location = st.text_input("Location", "")
    epa = st.text_input("EPA (Optional)", "")
    ta = st.text_input("T/A (Optional)", "")

    # Optional checkbox fields for location, EPA, and T/A
    include_location = st.checkbox("Include Location")
    include_epa = st.checkbox("Include EPA")
    include_ta = st.checkbox("Include T/A")

    # Apply filters for monthly report generation (optional)
    st.write("### Filter Tasks by Month 🔍")
    month_filter = st.date_input("Select the month:", datetime.today())
    filtered_df = df[df['Deadline'].str.contains(month_filter.strftime("%Y-%m"))]

    st.write(f"Filtered tasks for {month_filter.strftime('%B %Y')}:")
    st.write(filtered_df)

    # Image upload feature
    st.write("### Upload Image for Monthly Report 📸")
    image_file = st.file_uploader("Upload an Image (Optional)", type=["jpg", "png", "jpeg"])

    # Display the uploaded image (optional)
    if image_file is not None:
        st.image(image_file, caption='Uploaded Image', use_column_width=True)
        image_data = image_file
    else:
        image_data = None

    # Add a button to download the Excel and Word reports
    if st.button('Download Excel 📥'):
        excel_data = to_excel(filtered_df, officer, location, epa, ta, image_data)
        st.download_button(
            label="Click to Download Excel 📊",
            data=excel_data,
            file_name=f'mwanza_irrigation_tracker_{month_filter.strftime("%Y%m")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )

    if st.button('Download Word 📄'):
        word_data = to_word(filtered_df, officer, location, epa, ta, image_data)
        st.download_button(
            label="Click to Download Word Report 📝",
            data=word_data,
            file_name=f'mwanza_irrigation_report_{month_filter.strftime("%Y%m")}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True
        )

if __name__ == '__main__':
    main()
