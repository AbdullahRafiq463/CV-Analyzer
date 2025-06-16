from fpdf import FPDF

def generate_report(filepath, results):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Title
    pdf.set_font("Arial", size=16, style="B")
    pdf.cell(200, 10, txt="CV Analysis Report", ln=True, align="C")
    pdf.ln(10)

    # Table Header
    pdf.set_font("Arial", size=12)
    pdf.cell(100, 10, txt="Candidate", border=1, align="C")
    pdf.cell(40, 10, txt="Score", border=1, align="C")
    pdf.ln()

    # Table Rows
    for candidate, score in results:
        pdf.cell(100, 10, txt=candidate, border=1)
        pdf.cell(40, 10, txt=str(round(score, 2)), border=1, align="C")
        pdf.ln()

    pdf.output(filepath)
