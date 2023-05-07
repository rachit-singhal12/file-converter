from flask import Flask,render_template,request,url_for,send_file
from docx2pdf import convert
from pdf2docx import Converter
from pdfminer.high_level import extract_text_to_fp
from io import StringIO
from pdfminer.pdfparser import PDFSyntaxError
import PyPDF2

app = Flask(__name__)

app = Flask(__name__, static_url_path='/static')

@app.route('/')
def index():
    return render_template("index.html")

#docx to pdf
@app.route('/dtp.html')
def docx_to_pdf():
    return render_template("dtp.html")

@app.route('/dtp.html/dtp_submit',methods = ['post'])
def dtp_submit():
    file = request.files['file']
    file.save('path/file.docx')
    convert(r"path/file.docx",r"path/file.pdf")
    file_path = "C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file.pdf"
    return send_file(file_path,as_attachment=True)    

#pdf to docx
@app.route('/ptd.html')
def pdf_to_docx():
    return render_template("ptd.html")

@app.route('/ptd.html/ptd_submit',methods = ['post'])
def ptd_submit():
    file = request.files['file']
    file.save('path/file1.pdf')
    cv = Converter("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file1.pdf")
    cv.convert("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file1.docx", start=0, end=None)
    cv.close()
    file_path = "C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file1.docx"
    return send_file(file_path,as_attachment=True)   

#pdf to html
@app.route('/pth.html')
def pdf_to_html():
    return render_template("pth.html")

@app.route('/ptd.html/pth_submit',methods = ['post'])
def pth_submit():
    file = request.files['file']
    file.save('path/file2.pdf')
    try:
        with open("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file2.html", 'wb') as html_file:
            extract_text_to_fp(open("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file2.pdf", 'rb'), html_file, output_type='html')
    except PDFSyntaxError as e:
        print(f"Error parsing PDF: {e}")
    file_path = "C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file2.html"
    return send_file(file_path,as_attachment=True)
                     
#pdf to rtf
@app.route('/ptr.html')
def pdf_to_rtf():
    return render_template("ptr.html")

@app.route('/ptr.html/ptr_submit',methods = ['post'])
def ptr_submit():
    file = request.files['file']
    file.save('path/file3.pdf')
    with open("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file3.pdf", 'rb') as pdf_file:
        reader = PyPDF2.PdfReader("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file3.pdf")
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    
    with open("C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file3.txt", 'w', encoding='utf-8') as rtf_file:
        rtf_file.write(text)
    file_path = "C:\\Users\\lenovo\\Desktop\\PROJECT\\path\\file3.txt"
    return send_file(file_path,as_attachment=True)



if __name__ == '__main__':
    app.run(debug =True,port=5000)
