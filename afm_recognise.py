import locale
import os
import ghostscript
import pytesseract as tess

tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
from PyPDF2 import PdfFileWriter, PdfFileReader
import openpyxl
import xlrd
import shutil
from shutil import copyfile

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path

# NEW FEATURES
try:
    from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
except ImportError:
    # REQUIRES Extra installation 'pip install pyPdf'
    from pyPdf import PdfFileReader, PdfFileWriter

import win32com.client as win32

# Checking if list has duplicates
def checkIfDuplicates_1(listOfElems):
    ''' Check if given list contains any duplicates '''
    if len(listOfElems) == len(set(listOfElems)):
        return False
    else:
        return True

# Get the index of duplicated items
def list_duplicates_of(seq,item):
    start_at = -1
    locs = []
    while True:
        try:
            loc = seq.index(item,start_at+1)
        except ValueError:
            break
        else:
            locs.append(loc)
            start_at = loc
    return locs

# Merging the PDFs
def mergePDFs(input_files_list, output_stream_str):
    merger = PdfFileMerger()

    for pdf in input_files_list:
        merger.append(pdf)

    merger.write(output_stream_str)
    merger.close()

# Reversing a list
def Reverse_List(lst):
    return [ele for ele in reversed(lst)]

# Check if file is .xls
def check_for_xls_file(csv_file_path):
    if '.xls' in csv_file_path:
        return True
    else:
        return False

# Convert .xls to .xlsx file
def convert_xls_to_xlsx(csv_file_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(csv_file_path)
    wb.SaveAs("C:\\temp_pdf_page_by_page\\" + "_CONVERTED_XLS", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()                                                                                              # FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    return "C:\\temp_pdf_page_by_page\\" + "_CONVERTED_XLS.xlsx"

# END OF NEW FEATURES

files_not_send = []


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location=''):
    email_sender = 'email_sender'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('email', 'password')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
        files_not_send.append(attachment_location)

    return True


def pdf2jpeg(pdf_input_path, jpeg_output_path):
    args = ["pef2jpeg",  # actual value doesn't matter
            "-dNOPAUSE",
            "-sDEVICE=jpeg",
            "-r144",
            "-sOutputFile=" + jpeg_output_path,
            pdf_input_path]

    encoding = locale.getpreferredencoding()
    args = [a.encode(encoding) for a in args]

    ghostscript.Ghostscript(*args)
    ghostscript.cleanup()


def main_program(pdf_file_path, csv_file_path, email_subj, email_body):
    pdf_pages_cnt = 0
    try:
        os.mkdir("C:\\temp_pdf_page_by_page")
    except OSError:
        pass

    inputpdf = PdfFileReader(open(pdf_file_path, "rb"))
    for i in range(inputpdf.numPages):
        pdf_pages_cnt = pdf_pages_cnt + 1
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        with open("C:\\temp_pdf_page_by_page\\document-page%s.pdf" % i, "wb") as outputStream:
            output.write(outputStream)

    for i in range(0, pdf_pages_cnt, 1):
        temp = "C:\\temp_pdf_page_by_page\\document-page%s.pdf" % i
        temp1 = "C:\\temp_pdf_page_by_page\\document-page%s.jpg" % i
        pdf2jpeg(
            temp,
            temp1
        )

    txt = []
    for i in range(0, pdf_pages_cnt, 1):
        text = tess.image_to_string(r"C:\\temp_pdf_page_by_page\\document-page%s.jpg" % i)
        txt.append(text.split(" "))

    # Check if file is xls
    if (check_for_xls_file(csv_file_path)):
        fname = convert_xls_to_xlsx(csv_file_path)
    else :
        fname = csv_file_path

    #fname = csv_file_path

    wb = openpyxl.load_workbook(fname)
    # ws = wb.get_sheet_by_name("Worksheet")
    ws = wb["Worksheet"]

    mylist = []
    raw_position_in_excel = []
    i = 0
    for cell in ws['Q']:
        i = i + 1
        if str(cell.value) != "None":
            print(cell.value)
            mylist.append(cell.value)
            raw_position_in_excel.append(i)

    cnt = 0
    final = []
    positions = []
    position_in_excel = []
    for i in range(1, len(mylist), 1):
        for j in range(0, len(txt), 1):
            for k in range(0, len(txt[j]), 1):
                if str(mylist[i]) in txt[j][k]:
                    final.append(mylist[i])
                    positions.append(j + 1)
                    position_in_excel.append(raw_position_in_excel[i])
                    cnt = cnt + 1

    emails = []
    sheet = xlrd.open_workbook(csv_file_path).sheet_by_name("Worksheet")
    len3 = len(position_in_excel)
    for i in range(0, len3, 1):
        emails.append(sheet.cell_value(position_in_excel[i] - 1, 8))

    print(cnt)

    email_subject_final = "email_subject"
    email_body_final = email_body

    # NEW FEATURES
    # CHECK IF THERE IS ANY DUPLICATE ENTRIES IN THE LIST
    if (checkIfDuplicates_1(emails)):
        # DUPLICATED ENTRIES FOUND IN THE LIST
        # Get duplicates list position
        index_of_paths = 0
        for item in emails:
            # GET DUPLICATES INDEX
            duplicated_paths = []
            duplicates_Index = list_duplicates_of(emails, item)
            email_to_send = emails[0]
            for duplicated_index in duplicates_Index:
                # GET DUPLICATES ACTUAL PATH
                duplicated_paths.append('C:/temp_pdf_page_by_page/document-page%s.pdf' % (positions[duplicated_index] - 1))

            # MERGE ALL PDFs TOGETHER
            mergePDFs(duplicated_paths, 'C:/temp_pdf_page_by_page/duplicated-document%s.pdf' % index_of_paths)
            i += 1

            # DELETE DUPLICATE ENTRIES FROM ALL LISTS
            for item in Reverse_List(duplicates_Index):
                final.pop(item)
                positions.pop(item)
                emails.pop(item)

            # SEND EMAIL
            send_email(email_to_send,
                       email_subject_final,
                       email_body_final,
                       'C:/temp_pdf_page_by_page/duplicated-document%s.pdf' % index_of_paths)
            index_of_paths += 1

        # AFTER FOR LOOP COMPLETION CHECK IF EVERY EMAIL HAS SENT
        # IF NOT, THEN THE FOLLOWING FOR LOOP WILL EXECUTE TO SEND THE REST OF THE EMAILS
        if len(emails) != 0:
            for i in range(0, len(emails), 1):
                send_email(emails[i],
                            email_subject_final,
                            email_body_final,
                            'C:/temp_pdf_page_by_page/document-page%s.pdf' % (positions[i] - 1))

    else:
        # THERE IS NO DUPLICATED ITEMS IN THE LIST
        for i in range(0, len(emails), 1):
            send_email(emails[i],
                       email_subject_final,
                       email_body_final,
                       'C:/temp_pdf_page_by_page/document-page%s.pdf' % (positions[i] - 1))
    #END OF NEW FEATURES

    print(files_not_send)
    if len(files_not_send) != 0:
        try:
            os.mkdir("C:\\INVOICES_NOT_SEND")
            output_error_file = open("C:\\INVOICES_NOT_SEND\\Email που δεν στάλθηκαν.txt", "w")
            output_error_file.write("Δεν κατάφερα να στείλω τα παρακάτω τιμολόγια.\n")
            for error_file in files_not_send:
                output_error_file.write(error_file + '\n')

            output_error_file.write("Στον φάκελο θα βρείτε και τα αρχεία που δεν κατάφερα να στείλω.\n")
            output_error_file.close()
            counter = 0
            for error_file in files_not_send:
                counter = counter + 1
                destination = "C:/INVOICES_NOT_SEND/document-page%s.pdf" % counter
                copyfile(error_file, destination)
        except:
            pass

    shutil.rmtree("C:\\temp_pdf_page_by_page")