from PIL import Image
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from flask.helpers import redirect
from fpdf import FPDF
import os
from werkzeug.exceptions import RequestEntityTooLarge
from aspose.words import Document , SaveFormat
app = Flask(__name__, template_folder='templets', static_folder='statics')
app.config['upload_folder'] = os.path.join(
    os.path.abspath(os.path.dirname(__file__)), 'statics', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ALLOWED_EXTENSIONS'] = [
    '.jpg', '.jpeg', '.png', '.gif', '.docx', '.doc', '.xlxs', '.xlsx',
    '.pptx', '.ppt', '.txt'
]
app.config['PDF_FOLDER'] = os.path.join(
    os.path.abspath(os.path.dirname(__file__)), 'statics', 'pdfs')


@app.route('/')
@app.route('/home')
def index():
  print(app.config['upload_folder'])
  return render_template('index.html')


@app.route('/uploads', methods=['POST'])
def upload():
  try:
    file = request.files['file']
    if file:
      extension = os.path.splitext(file.filename)[1].lower()
      if extension not in app.config['ALLOWED_EXTENSIONS']:
        return 'File is not an image.'
      file.save(
          os.path.join(app.config['upload_folder'],
                       secure_filename(file.filename)))
      if extension == '.jpg' or extension == '.jpeg' or extension == '.png' or extension == '.gif':
        nameoffile = os.path.splitext(file.filename)[0]
        pdf_path = os.path.join(app.config['PDF_FOLDER'], f"{nameoffile}.pdf")
        image = Image.open(
            os.path.join(app.config['upload_folder'],
                         secure_filename(file.filename)))
        image.save(pdf_path, "PDF")
        return send_file(pdf_path,as_attachment=True,download_name=f"{nameoffile}.pdf")

      if extension == '.docx' or extension == '.doc':
        
        nameoffile = os.path.splitext(file.filename)[0]
        pdf_path = os.path.join(app.config['PDF_FOLDER'], f"{nameoffile}.pdf")
        doc = Document(os.path.join(app.config['upload_folder'],secure_filename(file.filename)))
        doc.save(pdf_path, SaveFormat.PDF)
        return send_file(pdf_path,
                         as_attachment=True,
                         download_name=f"{nameoffile}.pdf")
      if extension == '.txt':
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial" , size= 15)
        nameoffile = os.path.splitext(file.filename)[0]        
        pdf_path = os.path.join(app.config['PDF_FOLDER'], f"{nameoffile}.pdf")
        doc = os.path.join(app.config['upload_folder'],secure_filename(file.filename))
        with open(doc, 'r') as f:
          for line in f:
            pdf.cell(200, 10, txt = line, ln = True, align = 'L')
        pdf.output(pdf_path)
        return send_file(
          pdf_path,
          as_attachment=True,
                         download_name=f"{nameoffile}.pdf"
        )
    return redirect('/home')
  except RequestEntityTooLarge:
    return 'File is larger then 16MBs. Or something went wrong'


if __name__ == '__main__':
  app.run(debug=True)
