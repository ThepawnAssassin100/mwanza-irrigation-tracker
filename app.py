import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
import datetime

# Title
st.title("Mwanza District Irrigation Department - Weekly Task Tracker")

# Load or initialize the weekly task data
if "tasks" not in st.session_state:
    st.session_state.tasks = pd.DataFrame(columns=[
        "Week", "Date", "Activity", "Location", "Planned Output",
        "Actual Output", "Status", "Remarks"
    ])

# Form to add a new task
with st.form("Add Task"):
    week = st.selectbox("Week Number", [f"Week {i}" for i in range(1, 6)])
    date = st.date_input("Date")
    activity = st.text_input("Activity")
    location = st.text_input("Location")
    planned_output = st.text_input("Planned Output")
    actual_output = st.text_input("Actual Output")
    status = st.selectbox("Status", ["Planned", "In Progress", "Completed"])
    remarks = st.text_input("Remarks")
    submit = st.form_submit_button("Add Task")

    if submit:
        new_task = {
            "Week": week,
            "Date": date,
            "Activity": activity,
            "Location": location,
            "Planned Output": planned_output,
            "Actual Output": actual_output,
            "Status": status,
            "Remarks": remarks
        }
        st.session_state.tasks = pd.concat([
            st.session_state.tasks, pd.DataFrame([new_task])
        ], ignore_index=True)

# Display current weekly tasks
grouped = st.session_state.tasks.groupby("Week")
for week, data in grouped:
    st.subheader(week)
    st.dataframe(data)

# Export to Word document
def create_word_report(df):
    doc = Document()
    doc.add_heading("Mwanza District Irrigation Department - Monthly Report", level=1)
    current_month = datetime.datetime.now().strftime("%B %Y")
    doc.add_paragraph(f"Report for {current_month}")

    for week, group in df.groupby("Week"):
        doc.add_heading(week, level=2)
        table = doc.add_table(rows=1, cols=len(group.columns))
        hdr_cells = table.rows[0].cells
        for i, column in enumerate(group.columns):
            hdr_cells[i].text = column

        for _, row in group.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Download Monthly Report (Word)"):
    word_file = create_word_report(st.session_state.tasks)
    st.download_button(
        label="Download Word Report",
        data=word_file,
        file_name="Mwanza_Irrigation_Monthly_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

