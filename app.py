import os
from flask import Flask, request, render_template_string
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
from datetime import datetime

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
                to_date=row["ToDate"]
            )
        return f"✅ Certificates generated in '{CERT_FOLDER}' folder."
    return render_template_string(HTML)


def generate_certificate(name, regno, course, from_date, to_date):
    from datetime import datetime
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from PIL import Image
    import pandas as pd

    output_file = os.path.join(CERT_FOLDER, f"{name.replace(' ', '_')}.pdf")
    width, height = A4

    c = canvas.Canvas(output_file, pagesize=A4)
    c.drawImage(ImageReader(BACKGROUND_IMAGE),
                0, 0, width=width, height=height)

    # Format dates
    today = datetime.today().strftime("%d %B %Y")
    from_date_fmt = pd.to_datetime(from_date).strftime("%d/%m/%Y")
    to_date_fmt = pd.to_datetime(to_date).strftime("%d/%m/%Y")

    # Header Top Right
    c.setFont("Helvetica", 10)
    c.drawRightString(width - 40, height - 100, "+91 8189898884")
    c.drawRightString(width - 40, height - 115,
                      "hr@by8labs.com")
    c.drawRightString(width - 40, height - 130,
                      "www.by8labs.com")

    # Date Right Side
    c.setFont("Helvetica", 12)
    c.drawRightString(width - 40, height - 200, today)

    # Title Centered
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 300,
                        "TO WHOMSOEVER IT MAY CONCERN")

    # Main Body Content
    c.setFont("Helvetica", 12)
    text = c.beginText(60, height - 330)
    text.setLeading(18)
    content = f"""
    This is to certify that Ms. {name}, bearing Register Number {regno}, currently 
    pursuing B.Sc in the Department of Computer Science Kalaignar Karunanidhi Govt Arts 
    College for Women, Pudukkottai, has successfully completed an internship at Inno BY8LABS 
    Solution Private Limited. The internship was in the domain of {course} and was carried 
    out over a duration of one week, from {from_date_fmt} to {to_date_fmt}.

    During the span, they proved to be punctual and reliable individuals. Their learning abilities
    are commendable, showing a quick grasp of new concepts. Feedback and evaluations 
    consistently highlighted their strong learning curve. Furthermore, their interpersonal and
    communication skills are excellent. We take this opportunity to wish them the very
    best in their future endeavors.
    """
    for line in content.strip().split("\n"):
        text.textLine(line.strip())
    c.drawText(text)

    # Footer Signature & Contact
    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, 120, "Best Regards,")
    c.drawString(60, 100, "Mrs. V. Karthiga, B.E.")
    c.setFont("Helvetica", 12)
    c.drawString(60, 85, "Head of Training,")
    c.drawString(60, 70, "BY8LABS Inc,")

    # Footer Address Lines
    address_lines = [
        "#5861, Santhanathapuram Puram, 7th street, Pudukkottai – 622001. | +91 8189898884.",
        "#08-82, Redhills, Singapore – 150069. | +65 81532542."
    ]
    c.setFont("Helvetica", 10)
    y = 40
    for line in address_lines:
        c.drawString(60, y, line)
        y -= 12

    c.save()


if __name__ == '__main__':
    app.run(debug=True)
