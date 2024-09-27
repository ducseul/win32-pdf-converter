import os, sys, argparse
from flask import Flask, request, jsonify, send_file
import fitz
# import win32com.client as win32
# import xlwings as excelConvert
from waitress import serve
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from office_converter import WordConverter, ExcelConverter
from pdf_signature_extract import SignatureExtract, SignatureDetails
import os
import time

# import tempfile


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['LOGS_FOLDER'] = 'log'
app.config['FILE_RETENTION_DAYS'] = 1

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['LOGS_FOLDER'], exist_ok=True)

def cout(message: str):
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    port = app.config.get('PORT', '')
    log_message = f"{timestamp}::{port} - {message}"
    
    print(log_message, file=sys.stdout)
    
    with open(os.path.join(app.config['LOGS_FOLDER'], "log.txt"), "a") as f:
        f.write(f"{log_message}\n")


def convert_to_pdf(file_path, output_path):
    cout(f'Convert {file_path} to {output_path}')
    ext = file_path.split('.')[-1].lower()
    if ext == 'pdf':
        # If the file is already a PDF, we just copy it to the output path
        os.copy(file_path, output_path)
    elif ext in ['doc', 'docx']:
        word = WordConverter()
        try:
            start_time = time.time()
            word.convert(file_path, output_path)
            end_time = time.time()
            cout(f"Conversion completed in {end_time - start_time:.2f} seconds.")
            return output_path
        finally:
            word.close()
    elif ext in ['xls', 'xlsx']:
        start_time = time.time()
        excel = ExcelConverter()
        try:
            excel.convert(file_path, output_path)
            excel.close()
            end_time = time.time()
            print(f"Conversion completed in {end_time - start_time:.2f} seconds.")
            return output_path
        finally:
            excel.close()
    else:
        raise ValueError('Unsupported file type')
    
# @app.route('/convert2', methods=['POST'])
# def convert_file2():
#     if 'file' not in request.files:
#         return "No file part", 400

#     file = request.files['file']
#     if file.filename == '':
#         return "No selected file", 400

#     if file and file.filename.endswith('.docx'):
#         with tempfile.TemporaryDirectory() as tempdir:
#             from docx2pdf import convert
#             input_path = os.path.join(tempdir, file.filename)
#             file.save(input_path)
            
#             output_path = os.path.splitext(input_path)[0] + ".pdf"

#             # Đo thời gian bắt đầu chuyển đổi
#             start_time = time.time()
            
#             convert(input_path, output_path)

#             # Đo thời gian kết thúc chuyển đổi
#             end_time = time.time()
#             processing_time = end_time - start_time
            
#             return send_file(output_path, as_attachment=True)

#     return "Invalid file format", 400

@app.route('/ping', methods=['GET', 'POST'])
def health_check():
    return jsonify({'message': 'Hi, I am fine'}), 200

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    from werkzeug.utils import secure_filename
    filename = secure_filename(file.filename)
    
        # Add a timestamp to the filename to make it unique
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename_with_timestamp = f"{timestamp}_{filename}"
    
    print(f'Convert file {filename_with_timestamp}')
    file_path = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], filename_with_timestamp))
    file.save(file_path)


    output_filename = f'converted_{timestamp}_{filename.rsplit(".", 1)[0]}.pdf'
    output_path = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], output_filename))

    try:
        convert_to_pdf(file_path, output_path)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        cout(f"Error processing file {file_path}: {e}")
        return jsonify({'error': 'File conversion failed'}), 500

    return send_file(output_path, as_attachment=True, mimetype='application/pdf')

@app.route('/extract-signature', methods=['POST'])
def extract_signature():
    if 'file' not in request.files:
        return jsonify({"message": "No file part", "filename": "", "signatures": []}), 400

    file = request.files['file']
    filename = file.filename

    if filename == '':
        return jsonify({"message": "No selected file", "filename": "", "signatures": []}), 400

    try:
        # Save the uploaded file temporarily
        from werkzeug.utils import secure_filename
        filename = secure_filename(file.filename)
        
            # Add a timestamp to the filename to make it unique
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        filename_with_timestamp = f"{timestamp}_{filename}"
        
        print(f'Extract signature from file {filename_with_timestamp}')
        file_path = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], filename_with_timestamp))
        file.save(file_path)
        
        # Open the PDF file
        signatures_raw = SignatureExtract().get_pdf_signatures(filename = file_path);
        signatures = []
        for signature in signatures_raw:
            sig = SignatureDetails()
            sig.digest_algorithm = signature.digest_algorithm
            sig.signature_algorithm = signature.signature_algorithm
            sig.content_type = signature.content_type
            sig.type=signature.type
            sig.signer_contact_info=signature.signer_contact_info
            sig.signer_location=signature.signer_location
            sig.signing_time=signature.signing_time
            sig.signature_type=signature.signature_type
            sig.signature_handler=signature.signature_handler
            
            certificate = signature.certificate
            sig.valid_from=certificate.validity.not_before
            sig.valid_to=certificate.validity.not_after
            issuer = certificate.issuer
            
            sig.issuer_country_name=issuer.country_name
            sig.issuer_organization_name=issuer.organization_name
            sig.issuer_common_name=issuer.common_name
            
            subject = signature.certificate.subject
            sig.subject_country_name=subject.country_name
            sig.subject_organization_name=subject.organization_name
            sig.subject_organizational_unit_name="subject.organizational_unit_name"
            sig.subject_common_name=subject.common_name
            sig.subject_locality_name=signature.locality_name

            signatures.append(sig)
        os.remove(file_path)

        if not signatures:
            return jsonify({"message": "No signatures found", "filename": filename, "signatures": []}), 400

        signature_details_dict = [sd.to_dict() for sd in signatures]    
        return jsonify({"message": "Signatures extracted", "filename": filename, "signatures": signature_details_dict})

    except Exception as e:
        return jsonify({"message": str(e), "filename": filename, "signatures": []}), 500

# @app.route('/convert', methods=['POST'])
# def convert_file():
#     if 'source' not in request.form or 'destination' not in request.form:
#         return jsonify({'error': 'Missing form-data fields'}), 400
    
#     # Retrieve form-data fields
#     source = request.form['source']
#     destination = request.form['destination']
#     try:
#         convert_to_pdf(source, destination)
#     except ValueError as e:
#         return jsonify({'error': str(e)}), 400
#     except Exception as e:
#         cout(f"Error processing file {source}: {e}")
#         return jsonify({'error': 'File conversion failed'}), 500

#     return destination, 200

def delete_old_files():
    cout('Cron:: Remove old file scanning')
    now = datetime.now()
    retention_period = timedelta(days=app.config['FILE_RETENTION_DAYS'])

    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path):
                file_modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                if now - file_modified_time > retention_period:
                    try:
                        os.remove(file_path)
                        cout(f"Deleted old file: {file_path}")
                    except Exception as e:
                        cout(f"Error deleting file {file_path}: {e}")

                        
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Run Flask app with specified port and debug mode")
    
    parser.add_argument('--keep_old_file', type=lambda x: (str(x).lower() == 'true'), default=True, help='Scan and remove old file after 1 day')
    parser.add_argument('--port', type=int, default=5000, help='Port to run the app on. Default 5000')
    parser.add_argument('--threads', type=int, default=4, help='Maximum threads being create. Default 4')
    parser.add_argument('--connection_limit', type=int, default=200, help='Maxium connection being serve. Default 200')
    parser.add_argument('--debug', type=lambda x: (str(x).lower() == 'true'), default=False, help='Run app in debug mode. Note that it will using Flask as backend. Some features might not available. Not recommend for running as product.')
    
    args = parser.parse_args()
    app.config['PORT'] = args.port
    
    cout(f'Start application with port {args.port}. Debug mode set to {args.debug}')
    if args.keep_old_file == True:
        cout('Remove old file is enable')
        scheduler = BackgroundScheduler()
        scheduler.add_job(delete_old_files, 'interval', days=1)
        scheduler.start()
        
    if args.debug:
        # Debug mode (not recommended for production)
        app.run(port=args.port, debug=args.debug, host='0.0.0.0', threaded=True)
    else:
        cout(f'Connection limit: {args.connection_limit}')
        cout(f'Threads limit: {args.threads}')
        cout(f'File uploaded being auto-remove after 1 days: {args.keep_old_file}')
        # Production mode with Waitress
        serve(app, host='0.0.0.0', port=args.port, connection_limit=args.connection_limit)
