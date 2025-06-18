import os
from flask import Flask, request, render_template_string
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
import pandas as pd
import os
from reportlab.lib.colors import orange, blue

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
CERT_FOLDER = 'certificates'
BACKGROUND_IMAGE = 'certificate.jpg'  # Make sure this file exists

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CERT_FOLDER, exist_ok=True)

HTML = '''
<h2>Upload Excel File to Generate Certificates</h2>
<form method="POST" enctype="multipart/form-data">
    <input type="file" name="excel" accept=".xlsx"><br><br>
    <input type="submit" value="Generate">
</form>
'''


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel']
        if not file:
            return "No file uploaded!"
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        for _, row in df.iterrows():
            generate_certificate(
                name=row["Name"],
                regno=row["RegisterNumber"],
                course=row["Course"],
                from_date=row["FromDate"],
                to_date=row["ToDate"],
                college=row["CollegeName"],
                department=row["Department"],
                degree=row["Degree"],
                duration=row["Duration"]
            )
        return f"✅ Certificates generated in '{CERT_FOLDER}' folder."
    return render_template_string(HTML)


def generate_certificate(name, regno, course, from_date, to_date, degree, department, college, duration):
    from datetime import datetime

    width, height = A4
    output_file = os.path.join("certificates", f"{name.replace(' ', '_')}.pdf")
    bg_img = ImageReader("certificate.jpg")

    from_date_fmt = pd.to_datetime(from_date).strftime("%d/%m/%Y")
    to_date_fmt = pd.to_datetime(to_date).strftime("%d/%m/%Y")
    today = datetime.today().strftime("%d %B %Y")

    c = canvas.Canvas(output_file, pagesize=A4)
    c.drawImage(bg_img, 0, 0, width=width, height=height)

    # Header info
    c.setFont("Helvetica", 12)
    c.drawRightString(width - 40, height - 100, "+91 8189898884")
    c.drawRightString(width - 40, height - 120,
                      "hr@by8labs.com")
    c.drawRightString(width - 40, height - 140,
                      "www.by8labs.com")

    c.setFont("Helvetica", 14)
    c.drawRightString(width - 40, height - 200, today)

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 280,
                        "TO WHOMSOEVER IT MAY CONCERN")

    # Paragraph Style
    styles = getSampleStyleSheet()
    para_style = ParagraphStyle(
        'Justify',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=18,
        alignment=TA_JUSTIFY
    )

    # Content
    content = f"""
    &nbsp;&nbsp;&nbsp;&nbsp;This is to certify that <b>{name}</b>, bearing Register Number <b>{regno}</b>, currently pursuing {degree} in the Department of {department} <b>{college}</b>, Pudukkottai, has successfully completed an internship at Inno <b>BY8LABS</b> Solution Private Limited. The internship was in the domain of {course} and was carried out over a duration of {duration}, from <b>{from_date_fmt} to {to_date_fmt}</b>.<br/>

    &nbsp;&nbsp;&nbsp;&nbsp;During the span, they proved to be punctual and reliable individuals. Their learning abilities are commendable, showing a quick grasp of new concepts. Feedback and evaluations consistently highlighted their strong learning curve. Furthermore, their interpersonal and communication skills are excellent. We take this opportunity to wish them the very best in their future endeavors.
    """

    paragraph = Paragraph(content.strip(), para_style)

    # Frame to hold the paragraph (centered block with margins)
    frame_margin_x = 60  # Left and right padding
    frame_width = width - 2 * frame_margin_x
    frame_height = 330
    frame_y = height - 650  # vertical starting point

    frame = Frame(frame_margin_x, frame_y, frame_width,
                  frame_height, showBoundary=0)
    frame.addFromList([paragraph], c)

    # Signature
    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, 220, "Best Regards,")
    signature_path = "nancy sign.png"  # use your actual file path
    if os.path.exists(signature_path):
        c.drawImage(signature_path, 60, 180, width=100, height=30, mask='auto')
    c.drawString(60, 165, "Mrs. S. Nancy MCA.,")
    c.setFont("Helvetica", 12)
    c.drawString(60, 150, "HR Manager,")
    c.drawString(60, 135, "BY8LABS Inc,")
    c.drawString(60, 120, "Pudukkottai - 622001")

    c.setFillColor(blue)
    c.rect(0, 55, width, 5, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)

    # Footer address
    footer = [
        "#5861, Santhanathapuram Puram, 7th street, Pudukkottai – 622001. | +91 8189898884.",
        "#08-82, Redhills, Singapore – 150069. | +65 81532542."
    ]
    c.setFont("Helvetica", 12)
    y = 40
    for line in footer:
        c.drawCentredString(width/2, y, line)
        y -= 14

    c.setFillColor(orange)
    c.rect(0, 0, width, 20, fill=1, stroke=0)

    c.save()


if __name__ == '__main__':
    app.run(debug=True)
