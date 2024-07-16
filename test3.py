from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def convert_to_pdf(file_path):
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"
    c = canvas.Canvas(pdf_path, pagesize=letter)
    with open(file_path, 'r') as file:
        lines = file.readlines()
        y = 750  # Starting y position
        for line in lines:
            c.drawString(50, y, line.strip())
            y -= 12  # Decrease y position for next line
            if y <= 50:  # Move to the next page if y reaches the bottom
                c.showPage()
                c = canvas.Canvas(pdf_path, pagesize=letter)
                y = 750  # Reset y position
        c.save()
    messagebox.showinfo("Success", "Text file converted to PDF successfully!")
