import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from datetime import datetime

# Function to create and export Word document
def to_word(filtered, officer, epa, ta):
    doc = Document()
    doc.add_heading('Mwanza District Irrigation Task Report 🌱📋', 0)

    doc.add_paragraph(f"🧑‍💼 Responsible Officer: {officer}")
    if epa:
        doc.add_paragraph(f"📍 EPA: {epa}")
    if ta:
        doc.add_paragraph(f"🏘️ T/A: {ta}")
    doc.add_paragraph(f"📅 Date: {datetime.now().strftime('%Y-%m-%d')}")

    doc.add_heading('📝 Task Details:', level=1)
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
    st.markdown("""
        <h1 style="color:#4CAF50; font-size:40px; text-align:center;">🌾 Mwanza District Irrigation Tracker 🌦️</h1>
        <p style="font-size:18px; text-align:center;">📌 Track field and office tasks 🛠️ with responsible officers, descriptions, and remarks. Export monthly reports in Word format. 📤</p>
    """, unsafe_allow_html=True)

    # Initial task data
    if 'tasks' not in st.session_state:
        st.session_state.tasks = pd.DataFrame(columns=[
            '📌 Task', '📝 Work Description', '📆 Deadline', '📊 Status', '🚧 Progress', 
            '🏢 Office Work', '💬 Remarks'
        ])

    st.subheader("➕ Add New Task")
    with st.form("task_form"):
        task = st.text_input("📌 Task")
        description = st.text_input("📝 Work Description")
        deadline = st.date_input("📆 Deadline")
        status = st.selectbox("📊 Status", ["In Progress 🚧", "Completed ✅", "Pending ⏳"])
        progress = st.text_input("🚧 Progress (e.g., 50%)")
        office_work = st.text_input("🏢 Office Work")
        remarks = st.text_input("💬 Remarks")
        submitted = st.form_submit_button("📥 Add Task")
        if submitted and task:
            new_row = pd.DataFrame.from_dict([{
                '📌 Task': task,
                '📝 Work Description': description,
                '📆 Deadline': deadline.strftime('%Y-%m-%d'),
                '📊 Status': status,
                '🚧 Progress': progress,
                '🏢 Office Work': office_work,
                '💬 Remarks': remarks
            }])
            st.session_state.tasks = pd.concat([st.session_state.tasks, new_row], ignore_index=True)

    st.subheader("📋 Current Tasks")
    if not st.session_state.tasks.empty:
        edited_df = st.data_editor(
            st.session_state.tasks,
            num_rows="dynamic",
            use_container_width=True,
            key="task_editor"
        )
        st.session_state.tasks = edited_df

    officer = st.text_input("🧑‍💼 Responsible Officer")
    epa = st.text_input("📍 EPA (optional)")
    ta = st.text_input("🏘️ T/A (optional)")
    selected_month = st.date_input("📅 Select Month", datetime.today())
    month_str = selected_month.strftime("%Y-%m")
    filtered_df = st.session_state.tasks[st.session_state.tasks['📆 Deadline'].str.startswith(month_str)]

    st.write(f"📅 Filtered tasks for {selected_month.strftime('%B %Y')}: 🗂️")
    st.dataframe(filtered_df)

    if st.button('📥 Download Word Report'):
        word_data = to_word(filtered_df, officer, epa, ta)
        st.download_button(
            label="📄 Download Word Document",
            data=word_data,
            file_name=f'mwanza_irrigation_report_{month_str}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

if __name__ == '__main__':
    main()
