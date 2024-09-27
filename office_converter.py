import platform
import os, sys, argparse
from datetime import datetime, timedelta

def cout(message: str):
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log_message = f"{timestamp} - {message}"
    
    print(log_message, file=sys.stdout)
    
    # with open(os.path.join(app.config['LOGS_FOLDER'], "log.txt"), "a") as f:
    #     f.write(f"{log_message}\n")

class WordConverter:

    
    def __init__(self):
        if platform.system() == "Windows":
            import pythoncom
            import win32com.client as win32

        pythoncom.CoInitialize()    
        # self.word = win32.DispatchEx('Word.Application')
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False
    def convert(self, file_path, output_path):
        try:
            if not self.word:
                raise ValueError("Word application is not initialized.")
            
            print(f"Word Application Object: {self.word}")
            try:
                doc = self.word.Documents.Open(
                    file_path,
                    ReadOnly=True,
                    NoEncodingDialog=True,
                    PasswordDocument="123" 
                )
                doc.Activate()
            except Exception as e:
                if "password" in str(e).lower():
                    raise ValueError("The file is password protected and cannot be opened.") from e
                else:
                    raise e
            doc = self.word.Documents.Open(file_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 is the code for PDF
            doc.Close()
        except Exception as e:
            print(f"Error converting {file_path} to PDF: {e.__class__.__name__}: {str(e)}")
            raise e
        finally:
            self.word.Quit()

    def close(self):
        if platform.system() == "Windows":
            import pythoncom
            import win32com.client as win32
        self.word.Quit()
        self.word = None
        pythoncom.CoUninitialize()

class ExcelConverter:
    def __init__(self):
        if platform.system() == "Windows":
            import pythoncom
            import win32com.client as win32
        pythoncom.CoInitialize()
        self.excel = win32.DispatchEx('Excel.Application')
        self.excel.Visible = False  # Set to True for debugging purposes if needed

    def convert(self, file_path, output_path):
        try:
            if not self.excel:
                raise ValueError("Excel application is not initialized.")
            
            print(f"Excel Application Object: {self.excel}")  # Debug: Print the Excel object
            
            workbook = self.excel.Workbooks.Open(file_path)
            workbook.ExportAsFixedFormat(0, output_path)  # 0 for PDF
            workbook.Close(False)
        except Exception as e:
            print(f"Error converting {file_path} to PDF: {e.__class__.__name__}: {str(e)}")
            raise e

    def close(self):
        try:
            if platform.system() == "Windows":
                import pythoncom
                import win32com.client as win32
            if self.excel:
                self.excel.Quit()
                self.excel = None
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Error closing Excel application: {e.__class__.__name__}: {str(e)}")