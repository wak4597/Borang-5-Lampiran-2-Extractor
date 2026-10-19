import io
import re
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table

st.set_page_config(page_title="Document Extractor", page_icon="📄", layout="centered")

st.title("📄 Borang 5 to Lampiran 2 Extractor")
st.write("Upload your Word documents (.docx) below to extract data into a 2-column Word table.")

def extract_field_data(doc):
    """Extracts target fields using structural parsing and regex."""
    data = {
        "nama_premis": "",
        "alamat_premis": "",
        "no_fail": "",
        "pegawai_proses": "",
        "isu_ulasan": ""
    }

# Extract text from regular paragraphs
    paragraph_lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # Extract text from tables
    table_lines = []
    for table in doc.tables:
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_cells:
                table_lines.append(" | ".join(row_cells))

    # Combine both into a single searchable text block
    full_document_text = "\n".join(paragraph_lines + table_lines)

    # 1. No. Fail
    fail_match = re.search(r'(KKM[\s/]*87/A2/[A-Za-z0-9/.-]+(?:[\s]*JILID[\s]*\d+)?)', full_document_text, re.IGNORECASE)
    if fail_match:
        data["no_fail"] = fail_match.group(1).strip()

    # 2. Nama Premis & Alamat Premis
    for table in doc.tables:
        for row in table.rows:
            row_text = " ".join([c.text for c in row.cells])
            
            if "Nama Premis:" in row_text:
                raw_premis = row.cells[-1].text.strip()
                if data["no_fail"]:
                    raw_premis = raw_premis.replace(data["no_fail"], "").strip()
                data["nama_premis"] = raw_premis.split('\n')[0].strip()

            if "Alamat Premis:" in row_text:
                data["alamat_premis"] = row.cells[-1].text.strip().replace('\n', ' ')

# 3. Pegawai Proses
    # Looking for either "Disediakan oleh" or "Pegawai proses:" followed by colons, optional pipes, and optional prefixes
    pegawai_match = re.search(
        r'(?:Disediakan\s*oleh|Pegawai\s*Proses)[\s:\|]*([^\n\r\|]+)',
        full_document_text, 
        re.IGNORECASE
    )
    if pegawai_match:
        data["pegawai_proses"] = pegawai_match.group(1).strip()

    # 4. Isu / Ulasan (UPDATED: Row-specific extraction & merged cell handler)
    for table in doc.tables:
        for row in table.rows:
            row_cells = []
            seen_cells = set()
            
            for cell in row.cells:
                # cell._tc is the underlying XML element. 
                # Merged cells share the exact same _tc object.
                if cell._tc not in seen_cells:
                    seen_cells.add(cell._tc)
                    clean_text = cell.text.strip()
                    if clean_text:
                        row_cells.append(clean_text)
            
            if row_cells:
                row_text = "\n".join(row_cells)
                
                # Only trigger the regex if the row contains specific keywords
                if "Ulasan Sekretariat:" in row_text or "Mencadangkan" in row_text:
                    # UPDATED REGEX: Captures from "Mencadangkan" until the end of the line containing "Borang 6"
                    ulasan_match = re.search(
                        r'(Mencadangkan[\s\S]*?Borang 6[^\n\r]*(?:\r?\n[^\n\r]+)*)',
                        row_text,
                        re.IGNORECASE
                    )
                    if ulasan_match:
                        data["isu_ulasan"] = ulasan_match.group(1).strip()
                        break # Stop searching this table once found
        
        if data["isu_ulasan"]:
            break # Stop searching other tables once found

    return data

# --- Formatting Helpers ---
BOLD_TERMS = [
    r"diperakui dengan terma atau syarat",
    r"diperakui"
]

ENGLISH_WORDS = {
    "accident", "activities", "adl", "administration", "administrator",
    "affiliated", "amp", "and", "angiografi", "area", "assessment",
    "association", "asthma", "avenue", "barrier", "based", "basement",
    "bay", "bays", "beam", "bed", "bedded", "bereavement", "block",
    "bobath", "brachytherapy", "breastfeeding", "bulk", "bunker",
    "business", "c-arm", "cafeteria", "care", "cashier", "catheterization",
    "centre", "chair", "changing", "chemotherapy", "children", "clean",
    "clinic", "clinical", "clinics", "closet", "cold", "collection",
    "community", "cone", "consumable", "corridor", "couch", "counselling",
    "counter", "crrt", "cssd", "cssu", "cytotoxic", "day", "daycare",
    "dayward", "dental", "department", "desk", "dialysis", "dining",
    "dirty", "disposal", "district", "doctor's", "double", "double-bedded",
    "drug", "east", "echo", "ecg", "eeg", "endoscopy", "entrance",
    "equipment", "examination", "exam", "executive", "existing",
    "extended", "eye", "facilities", "facillities", "female", "fertility",
    "filing", "first", "floor", "focused", "four",
    "four-bedded", "function", "gamma", "general", "ground", "gym",
    "gymnasium", "haemodialysis", "hand", "hdu", "health", "high",
    "holding", "home", "hospice", "hospital", "hot", "icu", "imaging",
    "inhouse", "intensity", "intermediate", "international", "intra",
    "isolation", "it", "janitor", "janitor's", "joint", "knife",
    "laboratory", "laser", "ldr", "level", "lift", "linear", "lobby",
    "lounge", "lower", "lung", "major", "male", "management", "manager",
    "maternity", "medical", "medication", "medicine", "meeting",
    "mezzanine", "minor", "mobile", "multidisiplin", "negative",
    "nephrologist", "nerve", "new", "non-hepatitis", "nurse", "nursery",
    "nursing", "observation", "occupational", "off", "office", "oncology",
    "open", "operation", "optometrist", "optometry", "oral", "orthopaedic",
    "outpatient", "outsourced", "palliative", "park", "patient", "pediatric",
    "pharmacy", "physical", "physio", "physiotherapy", "physician's",
    "plastic", "play", "positive", "prep", "preparation", "pressure",
    "private", "procedure", "processing", "public", "quarantine",
    "radiology", "radiologist", "reception", "reconstitution", "record",
    "recovery", "refuge", "registration", "rehab", "rehabilitation",
    "relokasi", "reporting", "resusitasi", "retail", "robotic", "room",
    "rooms", "sample", "satellite", "scanner", "scn", "scope", "screening",
    "second", "serum", "server", "services", "settlement", "single",
    "single-bedded", "sleep", "sluice", "society", "soiled", "special",
    "specialist", "specialists", "speech", "staff", "station",
    "sterilization", "storage", "store", "stretcher", "study", "suite",
    "support", "surgery", "surgical", "table", "tea", "technical", "test",
    "theatre", "therapy", "third", "three", "three-bedded", "toilet",
    "tower", "traction", "training", "treadmill", "treatment", "triage",
    "trolley", "ultrasound", "unit", "utility", "vacant", "vascular",
    "vertigo", "vestibular", "vision", "visiting", "volunteer", "waiting",
    "ward", "wheelchair", "wing", "women's", "x-ray"
}

def format_styled_text_run(paragraph, text):
    """Formats right-column text with Arial 9pt, selective bolding, and italics for English words."""
    if not text:
        return

    bold_pattern = f"({'|'.join(BOLD_TERMS)})"
    tokens = re.split(bold_pattern, text, flags=re.IGNORECASE)

    for token in tokens:
        if not token:
            continue
            
        is_bold = any(re.fullmatch(term, token, re.IGNORECASE) for term in BOLD_TERMS)
        sub_tokens = re.split(r'(\s+|[.,;:]+)', token)
        
        for sub in sub_tokens:
            if not sub:
                continue
            
            run = paragraph.add_run(sub)
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            
            if is_bold:
                run.bold = True
            
            clean_word = re.sub(r'[^\w]', '', sub).lower()
            if clean_word in ENGLISH_WORDS and not is_bold:
                run.italic = True


def build_summary_doc(records):
    """Creates in-memory Word Document formatted with Arial 9pt, selective bold labels, and styling."""
    doc = Document()
    
    # Title Section
    title_p = doc.add_paragraph()
    title_run = title_p.add_run('Laporan Semakan Permohonan - Summarized Table')
    title_run.font.name = 'Arial'
    title_run.font.size = Pt(12)
    title_run.bold = True

    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    
    # Header Row
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Maklumat Premis & Pegawai", "Isu / Ulasan"]
    
    for i, title in enumerate(hdr_titles):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(9)
        run.bold = True

    # Populate Data
    for data in records:
        row_cells = table.add_row().cells
        
        # --- Left Column ---
        p_left = row_cells[0].paragraphs[0]
        
        # 1. Nama Premis (Value only, no label)
        val_premis = p_left.add_run(f"{data.get('nama_premis', '')}\n")
        val_premis.font.name = 'Arial'
        val_premis.font.size = Pt(9)
        val_premis.bold = True

        # 2. Alamat Premis (Value only, no label)
        val_alamat = p_left.add_run(f"{data.get('alamat_premis', '')}\n")
        val_alamat.font.name = 'Arial'
        val_alamat.font.size = Pt(9)
        val_alamat.bold = False

        # 3. No. Fail (Label + Value)
        lbl_fail = p_left.add_run("No. Fail: ")
        lbl_fail.font.name = 'Arial'
        lbl_fail.font.size = Pt(9)
        lbl_fail.bold = True
        
        val_fail = p_left.add_run(f"{data.get('no_fail', '')}\n")
        val_fail.font.name = 'Arial'
        val_fail.font.size = Pt(9)
        val_fail.bold = False

        # 4. Pegawai Proses (Label + Value)
        lbl_pegawai = p_left.add_run("Pegawai Proses: ")
        lbl_pegawai.font.name = 'Arial'
        lbl_pegawai.font.size = Pt(9)
        lbl_pegawai.bold = True
        
        val_pegawai = p_left.add_run(f"{data.get('pegawai_proses', '')}")
        val_pegawai.font.name = 'Arial'
        val_pegawai.font.size = Pt(9)
        val_pegawai.bold = False

        # --- Right Column ---
        p_right = row_cells[1].paragraphs[0]
        right_text = data.get('isu_ulasan', '')
        format_styled_text_run(p_right, right_text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- Streamlit UI ---
uploaded_files = st.file_uploader(
    "Choose Word files (only .docx!)",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Process Documents", type="primary"):
        with st.spinner("Extracting data from files..."):
            extracted_records = []
            
            for file in uploaded_files:
                doc = Document(file)
                record = extract_field_data(doc)
                extracted_records.append(record)

            output_buffer = build_summary_doc(extracted_records)

        st.success("Extraction Complete!")
        
        st.download_button(
            label="📥 Download Extracted Summary (.docx)",
            data=output_buffer,
            file_name="Lampiran 2 Borang 5 raw.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
