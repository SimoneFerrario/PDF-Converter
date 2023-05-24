import os
import win32com.client
from tqdm import tqdm

# Funzioni per la conversione
def doc_to_pdf(doc_path, pdf_path):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

def xls_to_pdf(xls_path, pdf_path):
    excel = win32com.client.Dispatch('Excel.Application')
    xls = excel.Workbooks.Open(xls_path)
    xls.ExportAsFixedFormat(0, pdf_path)
    xls.Close()
    excel.Quit()

def vsd_to_pdf(vsd_path, pdf_path):
    visio = win32com.client.Dispatch('Visio.InvisibleApp')
    doc = visio.Documents.Open(vsd_path)
    doc.ExportAsFixedFormat(1, pdf_path, 0, 0)
    doc.Close()
    visio.Quit()

# Chiedi all'utente il percorso della directory principale
directory_path = input("Inserisci il percorso della directory principale: ")

# Naviga nelle directory e converti i file
for dirpath, dirs, files in os.walk(directory_path):
    print(f"Processing directory: {dirpath}")
    for file in tqdm(files, desc="Converting files"):
        file_path = os.path.join(dirpath, file)
        print(f"Processing file: {file_path}")
        if file_path.endswith('.pdf'):
            continue
        elif file_path.endswith('.doc') or file_path.endswith('.docx'):
            pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
            doc_to_pdf(file_path, pdf_path)
        elif file_path.endswith('.xls') or file_path.endswith('.xlsx') or file_path.endswith('.csv'):
            pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
            xls_to_pdf(file_path, pdf_path)
        elif file_path.endswith('.vsd') or file_path.endswith('.vsdx'):
            pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
            vsd_to_pdf(file_path, pdf_path)
