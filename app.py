
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, send_from_directory, session
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from datetime import datetime
import uuid
import json

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_FOLDER'] = 'generated_docs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

# Store generated files info (in production, use database)
generated_files_store = {}

# ---------------------- Certificate Generator ----------------------
def generate_certificates(excel_file, template_file, output_folder, font_path="arialbd.ttf", font_size=100):
    """Generate certificates with perfectly centered names."""
    try:
        data = pd.read_excel(excel_file)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Try to load font, fallback to default if not available
        try:
            font_name = ImageFont.truetype(font_path, font_size)
        except:
            # Use default font
            font_name = ImageFont.load_default()

        generated_files = []
        for _, row in data.iterrows():
            name = str(row["Name"]).strip()
            certificate = Image.open(template_file)
            draw = ImageDraw.Draw(certificate)

            # Center the name text
            bbox = draw.textbbox((0, 0), name, font=font_name)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            x = (certificate.width - text_width) / 2
            y = 600  # Adjust vertically to fit your design

            draw.text((x, y), name, fill="orange", font=font_name)

            output_filename = f"certificate_{name.replace(' ', '_')}_{uuid.uuid4().hex[:8]}.png"
            output_path = os.path.join(output_folder, output_filename)
            certificate.save(output_path)
            generated_files.append({
                'name': name,
                'filename': output_filename,
                'type': 'certificate',
                'format': 'png',
                'path': output_path
            })

        return True, generated_files
    except Exception as e:
        return False, str(e)

# ---------------------- Associate Degree Functions ----------------------
def AssociateExcel_data(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def AssociateDocument(template, output_directory, student):
    doc = DocxTemplate(template)
    current_date = datetime.now().strftime("%B %d, %Y")
    doc.render({
        'name_kh': student[2],
        'g1': student[4],
        'id_kh': student[0],
        'name_e': student[3],
        'g2': student[5],
        'id_e': student[1],
        'dob_kh': student[6],
        'pro_kh': student[8],
        'dob_e': student[7],
        'pro_e': student[9],
        'ed_kh': student[10],
        'ed_e': student[11],
        'cur_date': current_date
    })
    filename_safe = student[3].replace(' ', '_').replace('/', '_')
    doc_name = os.path.join(output_directory, f"associate_{filename_safe}_{uuid.uuid4().hex[:8]}.docx")
    doc.save(doc_name)
    return doc_name

def AssociateConvertPDF(doc_path, pdf_directory):
    pdf_filename = os.path.splitext(os.path.basename(doc_path))[0] + ".pdf"
    pdf_path = os.path.join(pdf_directory, pdf_filename)
    convert(doc_path, pdf_path)
    return pdf_path

def generate_associate_documents(excel_file, template_file, option):
    """Generate associate documents with file tracking"""
    docx_directory = os.path.join(app.config['GENERATED_FOLDER'], "Associate_Documents")
    pdf_directory = os.path.join(app.config['GENERATED_FOLDER'], "Associate_PDF")

    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)
    data_rows = AssociateExcel_data(excel_file)

    generated_files = []
    for row in data_rows[1:]:
        if row[3]:  # Check if English name exists
            if option in ["doc", "both"]:
                doc_path = AssociateDocument(template_file, docx_directory, row)
                generated_files.append({
                    'name': row[3],
                    'filename': os.path.basename(doc_path),
                    'type': 'associate',
                    'format': 'docx',
                    'path': doc_path
                })

            if option in ["pdf", "both"]:
                if option == "pdf":
                    doc_path = AssociateDocument(template_file, pdf_directory, row)
                else:
                    doc_path = AssociateDocument(template_file, docx_directory, row)
                
                pdf_path = AssociateConvertPDF(doc_path, pdf_directory)
                generated_files.append({
                    'name': row[3],
                    'filename': os.path.basename(pdf_path),
                    'type': 'associate',
                    'format': 'pdf',
                    'path': pdf_path
                })
                
                if option == "pdf":
                    os.remove(doc_path)

    return True, generated_files

# ---------------------- Transcript Functions ----------------------
def TranscriptExcel_data(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def TranscriptDocument(template, output_directory, row_data):
    doc = DocxTemplate(template)
    current_date = datetime.now().strftime("%B %d, %Y")
    doc.render({
        "student_id": row_data[0],
        "first_name": row_data[1],
        "last_name": row_data[2],
        "logic": row_data[3],
        "l_g": row_data[4],
        "bcum": row_data[5],
        "bc_g": row_data[6],
        "design": row_data[7],
        "d_g": row_data[8],
        "p1": row_data[9],
        "p1_g": row_data[10],
        "e1": row_data[11],
        "e1_g": row_data[12],
        "wd": row_data[13],
        "wd_g": row_data[14],
        "algo": row_data[15],
        "al_g": row_data[16],
        "p2": row_data[17],
        "p2_g": row_data[18],
        "e2": row_data[19],
        "e2_g": row_data[20],
        "sd": row_data[21],
        "sd_g": row_data[22],
        "js": row_data[23],
        "js_g": row_data[24],
        "php": row_data[25],
        "ph_g": row_data[26],
        "db": row_data[27],
        "db_g": row_data[28],
        "vc1": row_data[29],
        "v1_g": row_data[30],
        "node": row_data[31],
        "no_g": row_data[32],
        "e3": row_data[33],
        "e3_g": row_data[34],
        "p3": row_data[35],
        "p3_g": row_data[36],
        "oop": row_data[37],
        "op_g": row_data[38],
        "lar": row_data[39],
        "lar_g": row_data[40],
        "vue": row_data[41],
        "vu_g": row_data[42],
        "vc2": row_data[43],
        "v2_g": row_data[44],
        "e4": row_data[45],
        "e4_g": row_data[46],
        "p4": row_data[47],
        "p4_g": row_data[48],
        "int": row_data[49],
        "in_g": row_data[50],
        'cur_date': current_date
    })
    filename_safe = f"{row_data[1]}_{row_data[2]}".replace(' ', '_').replace('/', '_')
    doc_name = os.path.join(output_directory, f"transcript_{filename_safe}_{uuid.uuid4().hex[:8]}.docx")
    doc.save(doc_name)
    return doc_name

def TranscriptPdf(doc_path, pdf_directory):
    pdf_filename = os.path.splitext(os.path.basename(doc_path))[0] + ".pdf"
    pdf_path = os.path.join(pdf_directory, pdf_filename)
    convert(doc_path, pdf_path)
    return pdf_path

def generate_transcripts(excel_file, template_file, option):
    """Generate transcripts with file tracking"""
    docx_directory = os.path.join(app.config['GENERATED_FOLDER'], "Transcript_Doc")
    pdf_directory = os.path.join(app.config['GENERATED_FOLDER'], "Transcript_PDF")

    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)
    data_rows = TranscriptExcel_data(excel_file)

    generated_files = []
    for row in data_rows[1:]:
        if row[1] and row[2]:  # Check if first and last name exist
            if option in ["doc", "both"]:
                doc_path = TranscriptDocument(template_file, docx_directory, row)
                generated_files.append({
                    'name': f"{row[1]} {row[2]}",
                    'filename': os.path.basename(doc_path),
                    'type': 'transcript',
                    'format': 'docx',
                    'path': doc_path
                })

            if option in ["pdf", "both"]:
                if option == "pdf":
                    doc_path = TranscriptDocument(template_file, pdf_directory, row)
                else:
                    doc_path = TranscriptDocument(template_file, docx_directory, row)
                    
                pdf_path = TranscriptPdf(doc_path, pdf_directory)
                generated_files.append({
                    'name': f"{row[1]} {row[2]}",
                    'filename': os.path.basename(pdf_path),
                    'type': 'transcript',
                    'format': 'pdf',
                    'path': pdf_path
                })
                
                if option == "pdf":
                    os.remove(doc_path)

    return True, generated_files

# ---------------------- Flask Routes ----------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        document_type = request.form.get('document_type')
        file_format = request.form.get('file_format', 'both')
        
        # Check if files were uploaded
        if 'excel_file' not in request.files or 'template_file' not in request.files:
            flash('Please upload both Excel and template files', 'error')
            return redirect(request.url)
        
        excel_file = request.files['excel_file']
        template_file = request.files['template_file']
        
        if excel_file.filename == '' or template_file.filename == '':
            flash('Please select both files', 'error')
            return redirect(request.url)
        
        # Validate file types
        if not (excel_file.filename.endswith('.xlsx') or excel_file.filename.endswith('.xls')):
            flash('Please upload a valid Excel file (.xlsx or .xls)', 'error')
            return redirect(request.url)
        
        # Save uploaded files
        excel_filename = f"{uuid.uuid4()}_{excel_file.filename}"
        template_filename = f"{uuid.uuid4()}_{template_file.filename}"
        
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        
        excel_file.save(excel_path)
        template_file.save(template_path)
        
        # Generate documents based on type
        try:
            generated_files = []
            if document_type == 'certificate':
                if not template_file.filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                    flash('Certificate generation requires a PNG or JPG template file', 'error')
                    return redirect(request.url)
                success, result = generate_certificates(
                    excel_path, 
                    template_path, 
                    os.path.join(app.config['GENERATED_FOLDER'], 'Certificates')
                )
                generated_files = result
            elif document_type == 'transcript':
                if not template_file.filename.lower().endswith('.docx'):
                    flash('Transcript generation requires a DOCX template file', 'error')
                    return redirect(request.url)
                success, result = generate_transcripts(excel_path, template_path, file_format)
                generated_files = result
            elif document_type == 'associate':
                if not template_file.filename.lower().endswith('.docx'):
                    flash('Associate document generation requires a DOCX template file', 'error')
                    return redirect(request.url)
                success, result = generate_associate_documents(excel_path, template_path, file_format)
                generated_files = result
            else:
                flash('Invalid document type', 'error')
                return redirect(request.url)
            
            if success:
                # Store files info in session
                session_id = str(uuid.uuid4())
                generated_files_store[session_id] = generated_files
                
                return redirect(url_for('results', 
                                      session_id=session_id,
                                      document_type=document_type,
                                      file_count=len(generated_files)))
            else:
                flash(f'Error generating documents: {result}', 'error')
                return redirect(request.url)
                
        except Exception as e:
            flash(f'Error processing files: {str(e)}', 'error')
            return redirect(request.url)
    
    return render_template('upload.html')

@app.route('/results')
def results():
    session_id = request.args.get('session_id')
    document_type = request.args.get('document_type', '')
    file_count = request.args.get('file_count', 0)
    
    if not session_id or session_id not in generated_files_store:
        flash('Session expired or invalid. Please generate documents again.', 'error')
        return redirect(url_for('upload'))
    
    files = generated_files_store[session_id]
    
    return render_template('results.html', 
                         document_type=document_type,
                         file_count=file_count,
                         files=files,
                         session_id=session_id)

@app.route('/download/<session_id>/<filename>')
def download_file(session_id, filename):
    """Download generated files"""
    if session_id not in generated_files_store:
        flash('Session expired', 'error')
        return redirect(url_for('upload'))
    
    # Find the file in our stored files
    files = generated_files_store[session_id]
    file_info = next((f for f in files if f['filename'] == filename), None)
    
    if not file_info or not os.path.exists(file_info['path']):
        flash('File not found', 'error')
        return redirect(url_for('results', session_id=session_id))
    
    return send_file(file_info['path'], as_attachment=True)

@app.route('/view/<session_id>/<filename>')
def view_file(session_id, filename):
    """View generated files in browser"""
    if session_id not in generated_files_store:
        flash('Session expired', 'error')
        return redirect(url_for('upload'))
    
    # Find the file in our stored files
    files = generated_files_store[session_id]
    file_info = next((f for f in files if f['filename'] == filename), None)
    
    if not file_info or not os.path.exists(file_info['path']):
        flash('File not found', 'error')
        return redirect(url_for('results', session_id=session_id))
    
    # For PDF and images, send file for viewing
    if file_info['format'] in ['pdf', 'png', 'jpg', 'jpeg']:
        return send_file(file_info['path'])
    else:
        # For DOCX, force download since browsers can't display them
        return send_file(file_info['path'], as_attachment=True)

@app.route('/batch_download/<session_id>')
def batch_download(session_id):
    """Download all files as zip"""
    if session_id not in generated_files_store:
        flash('Session expired', 'error')
        return redirect(url_for('upload'))
    
    # Simple implementation - redirect to first file
    files = generated_files_store[session_id]
    if files:
        return redirect(url_for('download_file', session_id=session_id, filename=files[0]['filename']))
    else:
        flash('No files to download', 'error')
        return redirect(url_for('results', session_id=session_id))

@app.route('/cleanup/<session_id>')
def cleanup(session_id):
    """Clean up generated files (optional)"""
    if session_id in generated_files_store:
        # In production, you might want to delete the physical files
        del generated_files_store[session_id]
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("🚀 Document Generator Server Starting...")
    print("📧 Open: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)