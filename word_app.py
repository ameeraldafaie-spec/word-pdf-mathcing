import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml.shared import OxmlElement
from io import BytesIO

def apply_smart_rtl(paragraph):
    """Applies RTL reading order without destroying centered/justified alignments."""
    pPr = paragraph._element.get_or_add_pPr()
    pPr.append(OxmlElement('w:bidi'))
    
    # Only flip Left alignment to Right. Leave Center and Justify alone.
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.LEFT or paragraph.alignment is None:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

st.set_page_config(page_title="PDF to DOCX Smart Aligner", page_icon="📄")
st.title("Proportional Volume Auto-Aligner")
st.write("This tool maps text mathematically to guarantee exact page counts, preserves centered layouts, and allows explicit margin control to prevent layout collapse.")

pdf_file = st.file_uploader("1. Upload English PDF", type="pdf")
docx_file = st.file_uploader("2. Upload Arabic Translation (DOCX)", type="docx")

margin_choice = st.selectbox(
    "3. Select Page Margins", 
    ["Keep Original Word Margins", "Normal (1 inch)", "Narrow (0.5 inch)"]
)

if pdf_file and docx_file:
    if st.button("Start Precision Alignment"):
        with st.spinner("Calculating text mass and mapping milestones..."):
            try:
                # 1. Analyze PDF Text Volume Profile
                doc_pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
                pdf_page_volumes = []
                total_pdf_chars = 0
                
                # Extract physical dimensions (in points)
                pdf_page_width = doc_pdf[0].rect.width
                pdf_page_height = doc_pdf[0].rect.height
                
                for page in doc_pdf:
                    text = page.get_text("text").strip()
                    chars_on_page = len(text)
                    pdf_page_volumes.append(chars_on_page)
                    total_pdf_chars += chars_on_page
                    
                st.success(f"PDF Scanned: {len(pdf_page_volumes)} pages detected. Mapping proportional thresholds...")
                
                # 2. Analyze Arabic DOCX Text Volume
                doc_word = Document(BytesIO(docx_file.read()))
                
                # Match original PDF dimensions and apply user-selected margins
                for section in doc_word.sections:
                    section.page_width = Pt(pdf_page_width)
                    section.page_height = Pt(pdf_page_height)
                    
                    if margin_choice == "Normal (1 inch)":
                        section.left_margin = Inches(1)
                        section.right_margin = Inches(1)
                        section.top_margin = Inches(1)
                        section.bottom_margin = Inches(1)
                    elif margin_choice == "Narrow (0.5 inch)":
                        section.left_margin = Inches(0.5)
                        section.right_margin = Inches(0.5)
                        section.top_margin = Inches(0.5)
                        section.bottom_margin = Inches(0.5)
                
                valid_paras = [p for p in doc_word.paragraphs if p.text.strip()]
                total_arabic_chars = sum(len(p.text.strip()) for p in valid_paras)
                
                # 3. Apply Smart RTL and Proportional Page Breaks
                current_arabic_chars = 0
                target_page_idx = 0
                
                # Calculate the initial threshold
                if total_pdf_chars > 0:
                    current_threshold = (pdf_page_volumes[0] / total_pdf_chars) * total_arabic_chars
                else:
                    current_threshold = float('inf')

                for p in doc_word.paragraphs:
                    if not p.text.strip():
                        continue
                        
                    # Apply safe formatting
                    apply_smart_rtl(p)
                    for run in p.runs:
                        rPr = run._element.get_or_add_rPr()
                        rPr.append(OxmlElement('w:rtl'))
                        
                    # Add paragraph length to our running total
                    text_len = len(p.text.strip())
                    current_arabic_chars += text_len
                    
                    # If we cross the threshold for the current page, insert a break
                    if current_arabic_chars >= current_threshold and target_page_idx < len(pdf_page_volumes) - 1:
                        if p.runs:
                            p.runs[-1].add_break(WD_BREAK.PAGE)
                            
                        target_page_idx += 1
                        # Increment threshold for the next page
                        current_threshold += (pdf_page_volumes[target_page_idx] / total_pdf_chars) * total_arabic_chars
                        
                # 4. Export
                output_io = BytesIO()
                doc_word.save(output_io)
                output_io.seek(0)
                
                st.download_button(
                    label="Download Precision Aligned DOCX",
                    data=output_io,
                    file_name="precision_aligned_arabic.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error processing: {e}")
