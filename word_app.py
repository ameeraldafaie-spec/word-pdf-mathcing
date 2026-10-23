import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_BREAK, WD_LINE_SPACING
from docx.shared import Pt
from docx.oxml.shared import OxmlElement
from io import BytesIO

def set_true_rtl(paragraph):
    """Forces actual Right-to-Left text direction in the Word XML."""
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    
    jc = OxmlElement('w:jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'right')
    pPr.append(jc)

st.set_page_config(page_title="PDF to DOCX Auto-Aligner", page_icon="📄")
st.title("English PDF to Arabic DOCX Auto-Aligner")
st.write("Upload the English PDF and the translated Arabic Word document. The app automatically detects original PDF dimensions and margins, compresses Arabic line spacing, maps page breaks, and enforces strict RTL formatting.")

pdf_file = st.file_uploader("1. Upload English PDF", type="pdf")
docx_file = st.file_uploader("2. Upload Arabic Translation (DOCX)", type="docx")

if pdf_file and docx_file:
    if st.button("Start Processing"):
        with st.spinner("Analyzing PDF margins and aligning documents..."):
            try:
                # 1. Analyze PDF Dimensions and Margins
                doc_pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
                paras_per_page = []
                
                pdf_page_width = doc_pdf[0].rect.width
                pdf_page_height = doc_pdf[0].rect.height
                
                # Variables to find the absolute bounding box of all text
                min_x0, min_y0 = float('inf'), float('inf')
                max_x1, max_y1 = 0, 0
                
                for page in doc_pdf:
                    blocks = page.get_text("blocks")
                    text_blocks = [b for b in blocks if b[4].strip()]
                    paras_per_page.append(len(text_blocks))
                    
                    # Deduce margins by checking coordinates (x0, y0, x1, y1)
                    for b in text_blocks:
                        x0, y0, x1, y1 = b[:4]
                        if x0 < min_x0: min_x0 = x0
                        if y0 < min_y0: min_y0 = y0
                        if x1 > max_x1: max_x1 = x1
                        if y1 > max_y1: max_y1 = y1
                
                # Calculate margins (adding a 15-point safety floor to avoid printer errors)
                left_margin = max(min_x0, 15)
                top_margin = max(min_y0, 15)
                right_margin = max(pdf_page_width - max_x1, 15)
                bottom_margin = max(pdf_page_height - max_y1, 15)
                
                st.success(f"PDF Analysis Complete: Found {len(paras_per_page)} pages. "
                           f"Detected Margins (pts) - L: {int(left_margin)}, R: {int(right_margin)}, "
                           f"T: {int(top_margin)}, B: {int(bottom_margin)}")
                
                # 2. Process DOCX
                doc_word = Document(BytesIO(docx_file.read()))
                
                # Apply the detected PDF dimensions and margins to the Word Document
                for section in doc_word.sections:
                    section.page_width = Pt(pdf_page_width)
                    section.page_height = Pt(pdf_page_height)
                    section.left_margin = Pt(left_margin)
                    section.right_margin = Pt(right_margin)
                    section.top_margin = Pt(top_margin)
                    section.bottom_margin = Pt(bottom_margin)

                valid_paras = [p for p in doc_word.paragraphs if p.text.strip()]
                
                for p in doc_word.paragraphs:
                    if p.text.strip():
                        set_true_rtl(p)
                        
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                        p.paragraph_format.line_spacing = 0.9 
                        
                        for run in p.runs:
                            rPr = run._element.get_or_add_rPr()
                            rtl = OxmlElement('w:rtl')
                            rPr.append(rtl)
                
                # 3. Inject Page Breaks
                current_idx = 0
                for i, count in enumerate(paras_per_page):
                    if i == len(paras_per_page) - 1:
                        break
                        
                    break_idx = current_idx + count - 1
                    
                    if break_idx < len(valid_paras):
                        if valid_paras[break_idx].runs:
                            valid_paras[break_idx].runs[-1].add_break(WD_BREAK.PAGE)
                            
                    current_idx += count
                    
                # 4. Save and Download
                output_io = BytesIO()
                doc_word.save(output_io)
                output_io.seek(0)
                
                st.download_button(
                    label="Download Aligned Arabic DOCX",
                    data=output_io,
                    file_name="aligned_arabic_translation.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"An error occurred: {e}")
