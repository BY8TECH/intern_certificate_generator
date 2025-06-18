import os
from flask import Flask, request, render_template_string, send_from_directory
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
import smtplib
from email.message import EmailMessage

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
CERT_FOLDER = 'certificates'
BACKGROUND_IMAGE = 'certificate.jpg'  # Make sure this file exists



EMAIL_SENDER = 'smby8labs@gmail.com'
EMAIL_PASSWORD = 'rdhz jjkr gnjr ehch'
EMAIL_SMTP = 'smtp.gmail.com'
EMAIL_PORT = 587

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CERT_FOLDER, exist_ok=True)

# HTML = '''
# <h2>Upload Excel File to Generate Certificates</h2>
# <form method="POST" enctype="multipart/form-data">
#     <input type="file" name="excel" accept=".xlsx"><br><br>
#     <input type="submit" value="Generate">
# </form>
# '''


HTML = '''
    <h2>Upload Excel File to Generate Certificates</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="excel" accept=".xlsx" required>
        <button type="submit">Generate</button>
    </form>
    <p><a href="/certificates">📄 View All Certificates</a></p>
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
        # return f"✅ Certificates generated in '{CERT_FOLDER}' folder."
        return f"✅ Certificates generated! <a href='/certificates'>View all</a>"
    return render_template_string(HTML)


# @app.route('/certificates')
# def list_certs():
#     files = os.listdir(CERT_FOLDER)
#     links = "".join(
#         f'<li><a href="/download/{file}" download>{file}</a></li>' for file in files
#     )
#     return f"<h2>Generated Certificates</h2><ul>{links}</ul>"



@app.route('/certificates')
def list_certs():
    files = sorted(f for f in os.listdir(CERT_FOLDER) if f.lower().endswith('.pdf'))
    return render_template_string("""
    <!DOCTYPE html>
    <html><head><title>Certificates</title>
    <style>
    body { font-family: sans-serif; margin: 40px; background: #f4f4f4; }
    #search { width: 60%; padding: 10px; margin: 20px auto; display: block; font-size: 16px; }
    ul { list-style-type: none; padding: 0; max-width: 700px; margin: auto; }
    li { background: white; margin: 10px 0; padding: 12px; border-radius: 5px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    input[type="email"] { padding: 4px; }
    button { padding: 4px 8px; }
    </style>
    <script>
    function filterCertificates() {
        const input = document.getElementById('search').value.toUpperCase();
        const items = document.querySelectorAll('li');
        items.forEach(li => {
            const text = li.innerText.toUpperCase();
            li.style.display = text.includes(input) ? '' : 'none';
        });
    }
    </script>
    </head><body>
    <h2 style="text-align:center;">Generated Certificates</h2>
    <input type="text" id="search" onkeyup="filterCertificates()" placeholder="Search...">
    <ul>
    {% for file in files %}
        <li>
            <a href="/download/{{ file }}" download>{{ file }}</a>
            <form method="POST" action="/send/{{ file }}">
                <input type="email" name="email" placeholder="Enter email" required />
                <button type="submit">Send</button>
            </form>
        </li>
    {% endfor %}
    </ul></body></html>
    """, files=files)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(CERT_FOLDER, filename, as_attachment=True)

@app.route('/send/<filename>', methods=['POST'])
def send_certificate(filename):
    email_to = request.form.get('email')
    filepath = os.path.join(CERT_FOLDER, filename)

    if not os.path.exists(filepath):
        return "❌ File not found", 404

    msg = EmailMessage()
    msg['Subject'] = 'Your Internship Certificate'
    msg['From'] = EMAIL_SENDER
    msg['To'] = email_to
    msg.set_content('Please find your certificate attached.')

    with open(filepath, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=filename)

    try:
        with smtplib.SMTP(EMAIL_SMTP, EMAIL_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        return f"✅ Email sent to {email_to}"
    except Exception as e:
        return f"❌ Email failed: {str(e)}", 500


# @app.route('/certificates')
# def list_certs():
#     # files = sorted(os.listdir(CERT_FOLDER))
#     files = files = sorted(f for f in os.listdir(CERT_FOLDER) if f.lower().endswith('.pdf'))
#     html = """
#     <!DOCTYPE html>
#     <html>
#     <head>
#         <title>Generated Certificates</title>
#         <style>
#             body {
#                 font-family: Arial, sans-serif;
#                 margin: 40px;
#                 background: #f4f4f4;
#             }
#             h2 {
#                 text-align: center;
#                 color: #333;
#             }
#             #search {
#                 width: 60%;
#                 padding: 10px;
#                 margin: 20px auto;
#                 display: block;
#                 font-size: 16px;
#             }
#             ul {
#                 list-style-type: none;
#                 padding: 0;
#                 max-width: 600px;
#                 margin: auto;
#             }
#             li {
#                 margin: 8px 0;
#                 background: white;
#                 padding: 12px 18px;
#                 border-radius: 5px;
#                 box-shadow: 0 2px 5px rgba(0,0,0,0.1);
#                 display: flex;
#                 justify-content: space-between;
#                 align-items: center;
#             }
#             a {
#                 text-decoration: none;
#                 color: #007BFF;
#             }
#             a:hover {
#                 text-decoration: underline;
#             }
#         </style>
#         <script>
#             function filterCertificates() {
#                 let input = document.getElementById('search');
#                 let filter = input.value.toUpperCase();
#                 let ul = document.getElementById('certList');
#                 let li = ul.getElementsByTagName('li');
#                 for (let i = 0; i < li.length; i++) {
#                     let a = li[i].getElementsByTagName("a")[0];
#                     let txt = a.textContent || a.innerText;
#                     li[i].style.display = txt.toUpperCase().includes(filter) ? "" : "none";
#                 }
#             }
#         </script>
#     </head>
#     <body>
#         <h2>Generated Certificates</h2>
#         <input type="text" id="search" onkeyup="filterCertificates()" placeholder="Search for names...">
#         <ul id="certList">
#             {% for file in files %}
#                 <li><a href="/download/{{ file }}" download>{{ file }}</a></li>
#             {% endfor %}
#         </ul>
#     </body>
#     </html>
#     """
#     return render_template_string(html, files=files)

# @app.route('/download/<filename>')
# def download_file(filename):
#     return send_from_directory(CERT_FOLDER, filename, as_attachment=True)


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
