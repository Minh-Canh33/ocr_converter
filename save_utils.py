from docx import Document
from fpdf import FPDF


# =========================
# SAVE DOCX
# =========================
def save_docx(path, text):

    doc = Document()

    doc.add_paragraph(text)

    doc.save(path)


# =========================
# SAVE PDF
# =========================
def save_pdf(path, text):

    pdf = FPDF()

    pdf.add_page()

    pdf.add_font(
        "DejaVu",
        "",
        "assets/DejaVuSans.ttf"
    )

    pdf.set_font("DejaVu", size=12)

    pdf.multi_cell(0, 10, text)

    pdf.output(path)