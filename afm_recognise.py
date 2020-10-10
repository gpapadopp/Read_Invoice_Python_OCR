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

files_not_send = []

def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location = ''):

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
    args = ["pef2jpeg", # actual value doesn't matter
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

    fname = csv_file_path


    wb = openpyxl.load_workbook(fname)
    ws = wb["Worksheet"]

    mylist = []
    raw_position_in_excel = []
    i = 0
    for cell in ws['Q']:
      i = i + 1
      if str(cell.value) != "None":
        print (cell.value)
        mylist.append(cell.value)
        raw_position_in_excel.append(i)

    cnt = 0
    final = []
    positions = []
    position_in_excel = []
    items = len(mylist)
    for i in range(1, items, 1):
      len1 = len(txt)
      for j in range(0, len1, 1):
        len2 = len(txt[j])
        for k in range(0, len2, 1):
          if str(mylist[i]) in txt[j][k]:
            final.append(mylist[i])
            positions.append(j+1)
            position_in_excel.append(raw_position_in_excel[i])
            cnt = cnt + 1

    emails = []
    loc = (csv_file_path)
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_name("Worksheet")
    len3 = len(position_in_excel)
    for i in range(0, len3, 1):
      emails.append(sheet.cell_value( position_in_excel[i]-1, 8))

    lenght_emails = len(emails)
    print(cnt)
    email_subject_final = "email_subject"
    email_body_final = email_body
    for i in range(0, lenght_emails, 1):
        send_email(emails[i],
                   email_subject_final,
                   email_body_final,
                   'C:/temp_pdf_page_by_page/document-page%s.pdf' % (positions[i]-1))


    shutil.rmtree("C:\\temp_pdf_page_by_page")
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