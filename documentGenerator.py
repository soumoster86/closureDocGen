import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from datetime import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------- PAGE CONFIG ---------------- #
st.set_page_config(page_title="Closure Document Generator", page_icon="📄", layout="wide")

# ---------------- STYLING ---------------- #
st.markdown("""
<style>
.block-container {padding-top: 2rem;}
h1, h2, h3 {color: #1f4e79;}
</style>
""", unsafe_allow_html=True)

# ---------------- SIDEBAR ---------------- #
with st.sidebar:
    st.header("📘 PM Guidance")
    st.info("""
✔ Keep inputs concise and structured  
✔ Focus on outcomes  
✔ Add evidence in deliverables  
✔ Avoid vague descriptions  
✔ Press F9 in Word to update TOC  
""")

# ---------------- HELPERS ---------------- #
def add_heading(doc, text, level):
    doc.add_heading(text, level=level)

def add_paragraph(doc, text):
    para = doc.add_paragraph(text if text else "N/A")
    para.paragraph_format.space_after = Pt(10)

def add_table(doc, data):
    table = doc.add_table(rows=1, cols=len(data[0]))
    for i, heading in enumerate(data[0]):
        table.rows[0].cells[i].text = heading
    for row in data[1:]:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val

def add_table_of_contents(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'

    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')

    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    run._element.append(fldChar_begin)
    run._element.append(instrText)
    run._element.append(fldChar_separate)
    run._element.append(fldChar_end)

def add_deliverables_section(doc, text, files):
    doc.add_heading("4. Deliverables", 1)
    add_paragraph(doc, text)

    if not files:
        return

    doc.add_paragraph("Attached Deliverables:")

    table_data = [["File Name", "Type"]]

    for file in files:
        file.seek(0)
        if file.type.startswith("image"):
            doc.add_paragraph(f"Image: {file.name}")
            doc.add_picture(BytesIO(file.read()), width=Inches(5))
        else:
            table_data.append([file.name, file.type])

    if len(table_data) > 1:
        add_table(doc, table_data)

def add_custom_sections(doc, sections):
    if not sections:
        return

    doc.add_heading("8. Additional Sections", 1)

    for i, sec in enumerate(sections, start=1):
        title = sec.get("title", "").strip()
        content = sec.get("content", "").strip()

        if title:
            doc.add_heading(f"{i}. {title}", 2)
            doc.add_paragraph(content if content else "N/A")

def generate_document(data):
    doc = Document()

    doc.add_heading(data["project_name"], 0)
    add_paragraph(doc, "Project Closure Report")
    add_paragraph(doc, f"Generated on: {datetime.now().strftime('%d %b %Y')}")

    doc.add_page_break()

    add_heading(doc, "Table of Contents", 1)
    add_table_of_contents(doc)
    doc.add_page_break()

    sections = [
        ("1. Project Overview", data["overview"]),
        ("2. Objectives", data["objectives"]),
        ("3. Timeline", data["timeline"]),
    ]

    for title, content in sections:
        add_heading(doc, title, 1)
        add_paragraph(doc, content)

    add_deliverables_section(doc, data["deliverables"], data["files"])

    remaining = [
        ("5. Risks & Issues", data["risks"]),
        ("6. Lessons Learned", data["lessons"]),
    ]

    for title, content in remaining:
        add_heading(doc, title, 1)
        add_paragraph(doc, content)

    add_heading(doc, "7. Stakeholders", 1)
    table = [["Name", "Role"]]
    table.extend(data["stakeholders"] or [["N/A", "N/A"]])
    add_table(doc, table)

    add_custom_sections(doc, data["custom_sections"])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------- UI ---------------- #
st.title("📄 Project Closure Generator")

# Inputs
project_name = st.text_input("Project Name", placeholder="E.g., Migration Project")
timeline = st.text_input("Timeline", placeholder="Jan 2025 – Sep 2025")

overview = st.text_area("Overview", placeholder="Summary...")
objectives = st.text_area("Objectives", placeholder="Goals...")
deliverables = st.text_area("Deliverables", placeholder="Outputs...")
uploaded_files = st.file_uploader("Upload Files", accept_multiple_files=True)

risks = st.text_area("Risks", placeholder="Risks...")
lessons = st.text_area("Lessons", placeholder="Lessons...")

# Custom Sections
if "custom_sections" not in st.session_state:
    st.session_state.custom_sections = []

if st.button("➕ Add Custom Section"):
    st.session_state.custom_sections.append({"title": "", "content": ""})

for i, sec in enumerate(st.session_state.custom_sections):
    st.session_state.custom_sections[i]["title"] = st.text_input(f"Title {i+1}", key=f"title_{i}")
    st.session_state.custom_sections[i]["content"] = st.text_area(f"Content {i+1}", key=f"content_{i}")

# Stakeholders
stakeholders = []
count = st.number_input("Stakeholders", 1, 10, 2)

for i in range(count):
    name = st.text_input(f"Name {i}", key=f"name_{i}")
    role = st.text_input(f"Role {i}", key=f"role_{i}")
    if name and role:
        stakeholders.append([name, role])

# ---------------- FULL PREVIEW ---------------- #
st.subheader("📄 Full Preview")

st.markdown(f"""
# {project_name or "Project Name"}

**Timeline:** {timeline or "N/A"}

## Overview
{overview or "N/A"}

## Objectives
{objectives or "N/A"}

## Deliverables
{deliverables or "N/A"}
""")

if uploaded_files:
    st.markdown("### 📎 Files")
    for f in uploaded_files:
        if f.type.startswith("image"):
            st.image(f)
        else:
            st.write(f.name)

st.markdown(f"""
## Risks
{risks or "N/A"}

## Lessons
{lessons or "N/A"}
""")

st.markdown("## Stakeholders")
for s in stakeholders:
    st.write(f"- {s[0]} ({s[1]})")

# Generate
if st.button("🚀 Generate Document"):
    data = {
        "project_name": project_name,
        "overview": overview,
        "objectives": objectives,
        "timeline": timeline,
        "deliverables": deliverables,
        "files": uploaded_files,
        "risks": risks,
        "lessons": lessons,
        "stakeholders": stakeholders,
        "custom_sections": st.session_state.custom_sections
    }

    doc = generate_document(data)

    st.download_button(
        "Download",
        doc,
        file_name="closure.docx"
    )