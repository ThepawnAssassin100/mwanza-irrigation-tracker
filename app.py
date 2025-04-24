import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from datetime import datetime

st.set_page_config(page_title="Mwanza Irrigation Tracker", page_icon="🌾", layout="centered")

st.title("🌾 Mwanza District Irrigation Department")
st.subheader("📅 Weekly Task Tracker")

if "tasks" not in st.session_state:
    st.session_state.tasks = pd.DataFrame(columns=["Week", "Task", "Responsible", "Status", "Start Date", "End Date", "Image"])

st.markdown("---")
st.markdown("### ➕ Add a New Task")

with st.form("task_form"):
    week = st.text_input("Week (e.g. Week 1 - April)")
    task = st.text_area("Task Description")
    responsible = st.text_input("Responsible Officer")
    start_date = st.date_input("Start Date")
    end_date = st.date_input("End Date")
    status = st.selectbox("Status", ["Planned", "In Progress", "Completed"])
    image = st.file_uploader("📷 Upload Task Photo (optional)", type=["jpg", "jpeg", "png"])
    submitted = st.form_submit_button("✅ Add Task")

if submitted:
    image_url = image.name if image else ""
    new_task = pd.DataFrame({
        "Week": [week], "Task": [task], "Responsible": [responsible],
        "Status": [status], "Start Date": [start_date], "End Date": [end_date], "Image": [image_url]
    })
    st.session_state.tasks = pd.concat([st.session_state.tasks, new_task], ignore_index=True)
    st.success("✅ Task added successfully")

st.markdown("---")
st.markdown("### 📋 Task List")

# Filters
week_filter = st.selectbox("🔎 Filter by Week", ["All"] + sorted(st.session_state.tasks["Week"].unique().tolist()))
status_filter = st.selectbox("🎯 Filter by Status", ["All", "Planned", "In Progress", "Completed"])

filtered = st.session_state.tasks
if week_filter != "All":
    filtered = filtered[filtered["Week"] == week_filter]
if status_filter != "All":
    filtered = filtered[filtered["Status"] == status_filter]

# Status color coding
status_icons = {
    "Planned": "🟡 Planned",
    "In Progress": "🟠 In Progress",
    "Completed": "🟢 Completed"
}
filtered["Status"] = filtered["Status"].map(status_icons)

st.dataframe(filtered.drop(columns=["Image"]), use_container_width=True)

# Optional image display
with st.expander("📸 View Task Photos"):
    for i, row in filtered.iterrows():
        if row["Image"]:
            st.markdown(f"**{row['Task']} ({row['Week']})**")
            st.image(row["Image"], caption=row["Task"], use_column_width=True)

# Summary
with st.expander("📊 Weekly Summary"):
    st.write("**Total Tasks:**", len(filtered))
    st.write("**Completed:**", (filtered["Status"] == "🟢 Completed").sum())

# Excel Export
def to_excel(df):
    output = BytesIO()
import pandas as pd

def to_excel(filtered):
    output = BytesIO()  # assuming you are using BytesIO or some output stream
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Sheet1')
        # No need for writer.save() because the context manager handles it automatically
    return output.getvalue()  # returning the Excel file content if needed

    processed_data = output.getvalue()
    return processed_data

excel_data = to_excel(filtered)
st.download_button("📊 Download Excel Report", excel_data, file_name="weekly_tasks.xlsx")

# Download Word report
def create_word_report(df):
    doc = Document()
    doc.add_heading("Mwanza District Irrigation - Monthly Report", 0)
    grouped = df.groupby("Week")
    for week, group in grouped:
        doc.add_heading(f"📅 {week}", level=1)
        for _, row in group.iterrows():
            doc.add_paragraph(f"📝 {row['Task']}", style='List Bullet')
            doc.add_paragraph(f"👤 Responsible: {row['Responsible']}")
            doc.add_paragraph(f"📌 Status: {row['Status']}")
            doc.add_paragraph(f"🗓️ Start: {row['Start Date']} → End: {row['End Date']}")
            if row['Image']:
                doc.add_paragraph(f"📷 Image: {row['Image']}")
            doc.add_paragraph("")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if not filtered.empty:
    docx_file = create_word_report(filtered)
    st.download_button("📥 Download Monthly Report (Word)", docx_file, file_name="mwanza_monthly_report.docx")
