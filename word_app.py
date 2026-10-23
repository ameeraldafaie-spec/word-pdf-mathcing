import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt
from docx.oxml.shared import OxmlElement
from io import BytesIO

def set_true_rtl(paragraph):
    """Forces actual Right-to-Left text direction in the Word XML."""
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    
    # Ensure the paragraph alignment defaults to right for RTL text
    jc = OxmlElement('w:jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'right')
    pPr.append(jc)

st.set_page_config(page_title="PDF to DOCX Auto-Aligner", page_icon="📄")
st.title("English PDF to Arabic DOCX Auto-Aligner")
st.write("Upload the English PDF and the translated Arabic Word document. The app will calculate the paragraph distribution in the PDF and mirror it in the Arabic document while preserving your formatting.")

pdf_file = st.file_uploader("1. Upload English PDF", type="pdf")
docx_file = st.file_uploader("2. Upload Arabic Translation (DOCX)", type="docx")

if pdf_file and docx_file:
    if st.button("Start Processing"):
        with st.spinner("Analyzing and aligning documents..."):
            try:
                # 1. Analyze PDF
                doc_pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
                paras_per_page = []
                
                for page in doc_pdf:
                    blocks = page.get_text("blocks")
                    # Filter out empty blocks/images
                    text_blocks = [b for b in blocks if b[4].strip()]
                    paras_per_page.append(len(text_blocks))
                    
                st.success(f"PDF Analysis Complete: Found {len(paras_per_page)} pages.")
                
                # 2. Process DOCX
                doc_word = Document(BytesIO(docx_file.read()))
                valid_paras = [p for p in doc_word.paragraphs if p.text.strip()]
                
                for p in doc_word.paragraphs:
                    if p.text.strip():
                        set_true_rtl(p)
                        # Iterate through runs to preserve inline styles while adjusting font size
                        for run in p.runs:
                            if run.font.size:
                                run.font.size = Pt(run.font.size.pt + 2)
                            else:
                                run.font.size = Pt(14)
                                
                            # RTL property for runs
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
                    
                # Save to BytesIO object for download
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
