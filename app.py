import streamlit as st
import pandas as pd
from io import BytesIO

# Function to create and export Excel file
def to_excel(filtered):
    # Create an in-memory BytesIO object to hold the Excel file
    output = BytesIO()

    # Use the context manager to write the DataFrame to Excel
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Sheet1')
        # No need for writer.save(), as it's handled by the context manager

    # Move the pointer back to the start of the BytesIO buffer
    output.seek(0)
    return output.getvalue()

# Streamlit interface
def main():
    st.set_page_config(page_title="Mwanza Irrigation Tracker", page_icon="🌾", layout="wide")
    
    # Title with Emoji and Styling
    st.markdown(
        """
        <h1 style="color:#4CAF50; font-size:40px; text-align:center;">Mwanza District Irrigation Tracker 🌿</h1>
        <p style="font-size:20px; text-align:center;">Track irrigation tasks and download task reports in Excel format.</p>
        """, unsafe_allow_html=True)

    # Example: Load some data (replace with your actual data loading process)
    data = {
        'Task': ['Irrigation Setup', 'Field Inspection', 'Maintenance'],
        'Deadline': ['2025-05-01', '2025-05-15', '2025-06-01'],
        'Status': ['In Progress', 'Completed', 'Pending']
    }

    df = pd.DataFrame(data)

    # Display data with styling
    st.write("### Irrigation Tasks List 📅")
    st.write(df.style.set_table_styles([
        {'selector': 'thead th', 'props': [('background-color', '#4CAF50'), ('color', 'white'), ('font-size', '14px')]},
        {'selector': 'tbody td', 'props': [('font-size', '12px')]},
    ]))

    # Add some spacing
    st.markdown("<hr>", unsafe_allow_html=True)

    # Button to download the data as an Excel file with an icon
    if st.button('Download Excel 📥', key="download_button"):
        excel_data = to_excel(df)
        st.download_button(
            label="Click to Download Excel 📊",
            data=excel_data,
            file_name='mwanza_irrigation_tracker.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )

    # Additional Section (Optional): Show completion stats
    st.markdown("<hr>", unsafe_allow_html=True)
    st.write("### Task Completion Overview 🔍")

    total_tasks = len(df)
    completed_tasks = len(df[df['Status'] == 'Completed'])
    in_progress_tasks = len(df[df['Status'] == 'In Progress'])
    pending_tasks = len(df[df['Status'] == 'Pending'])

    st.write(f"📈 **Total Tasks**: {total_tasks}")
    st.write(f"✅ **Completed**: {completed_tasks}")
    st.write(f"🔄 **In Progress**: {in_progress_tasks}")
    st.write(f"🕒 **Pending**: {pending_tasks}")

if __name__ == '__main__':
    main()
