from tkinter import *
from tkinter import ttk, filedialog, simpledialog
from tkinter.font import Font
import xlrd as xl
from afm_recognise import *
from tkinter import messagebox as mb
import os

dir_path = os.path.dirname(os.path.realpath(__file__))

window = Tk()

def choose_csv_file():
    global csv_file_name
    csv_file_name = filedialog.askopenfilename(initialdir="/", title="Επιλογή Πελατολογίου", filetypes=(("Excel Files (.xlsx)", "*.xlsx"), ("All Files", "*.*")))
    csv_textbox.config(state="normal")
    csv_textbox.delete('1.0', END)
    csv_textbox.insert(INSERT, csv_file_name)
    csv_textbox.config(state="disabled")
    wb = xl.open_workbook(csv_file_name)
    s1 = wb.sheet_by_index(0)
    s1.cell_value(0, 0)
    csv_counter.config(state="normal")
    csv_counter.insert(INSERT, s1.nrows)
    csv_counter.config(state="disabled")

def choose_pdf_file():
    global pdf_file_name
    pdf_file_name = filedialog.askopenfilename(initialdir="/", title="Επιλογή Τιμολογίων", filetypes=(("PDF Files (.pdf)", "*.pdf"), ("All Files", "*.*")))
    pdf_textbox.config(state="normal")
    pdf_textbox.delete('1.0', END)
    pdf_textbox.insert(INSERT, pdf_file_name)
    pdf_textbox.config(state="disabled")
    with open(pdf_file_name, "rb") as pdf_file:
        pdf_reader = PdfFileReader(pdf_file)
        pdf_counter.config(state="normal")
        pdf_counter.insert(INSERT, pdf_reader.numPages)
        pdf_counter.config(state="disabled")

def send_btn():
    subject = simpledialog.askstring("ΤΟ ΘΕΜΑ ΤΟΥ e-mail.", "ΠΑΡΑΚΑΛΩ ΠΛΗΚΤΡΟΛΟΓΗΣΤΕ ΤΟ ΘΕΜΑ ΤΟΥ e-mail.")
    body = simpledialog.askstring("ΤΟ ΚΥΡΙΩΣ ΚΕΙΜΕΝΟ ΤΟΥ e-mail.",
                                  "ΠΑΡΑΚΑΛΩ ΠΛΗΚΤΡΟΛΟΓΗΣΤΕ ΤΟ ΚΥΡΙΩΣ ΚΕΙΜΕΝΟ ΤΟΥ e-mail.")
    main_program(pdf_file_name, csv_file_name, subject, body)
    mb.showinfo('Επιτυχία', 'Το πρόγραμμα ολοκληρώθηκε με επιτυχία.')

window.title("ΑΠΟΣΤΟΛΗ ΤΙΜΟΛΟΓΙΩΝ")

canvas = Canvas(window, width=525, height=150)
canvas.pack()

separator = ttk.Separator(window, orient='horizontal')
separator.pack(side='top', fill='x')

choose_text = canvas.create_text(260, 140, text="Eπιλογή Αρχείων", fill="#652828", anchor=CENTER)

myFont = Font(family="Futura", size=11)
csv_textbox = Text(window, height=1.25, width=132, bg="light yellow")
csv_textbox.place(rely=0.35, relx=0.02, x=0, y=0, anchor=W)
csv_textbox.insert(INSERT, "Επιλογή Αρχείου")
csv_textbox.configure(font=myFont)
csv_textbox.config(state="disabled")

choose_csv_btn = Button(window, text ="Επιλογή Πελατολογίου", command=choose_csv_file)
choose_csv_btn.place(rely=0.35, relx=0.915, x=0, y=0, anchor=CENTER)

pdf_textbox = Text(window, height=1.25, width=132, bg="light yellow")
pdf_textbox.place(rely=0.50, relx=0.02, x=0, y=0, anchor=W)
pdf_textbox.insert(INSERT, "Επιλογή Αρχείου")
pdf_textbox.configure(font=myFont)
pdf_textbox.config(state="disabled")

choose_pdf_btn = Button(window, text ="  Επιλογή Τιμολογίων  ", command=choose_pdf_file)
choose_pdf_btn.place(rely=0.50, relx=0.915, x=0, y=0, anchor=CENTER)

csv_counter_text_box = Text(window, height=1.25, width=19, bg="white")
csv_counter_text_box.place(rely=0.6, relx=0.77, x=0, y=0, anchor=W)
csv_counter_text_box.insert(INSERT, "Επιλεγμένοι Πελάτες: ")
csv_counter_text_box.configure(font=myFont)
csv_counter_text_box.config(state="disabled")

csv_counter = Text(window, height=1.25, width=10, bg="white")
csv_counter.place(rely=0.6, relx=0.97, x=0, y=0, anchor=E)
csv_counter.insert(INSERT, "")
csv_counter.configure(font=myFont)
csv_counter.config(state="disabled")

pdf_counter_text_box = Text(window, height=1.25, width=19, bg="white")
pdf_counter_text_box.place(rely=0.7, relx=0.77, x=0, y=0, anchor=W)
pdf_counter_text_box.insert(INSERT, "Επιλεγμένα Τιμολόγια:")
pdf_counter_text_box.configure(font=myFont)
pdf_counter_text_box.config(state="disabled")

pdf_counter = Text(window, height=1.25, width=10, bg="white")
pdf_counter.place(rely=0.7, relx=0.97, x=0, y=0, anchor=E)
pdf_counter.insert(INSERT, "")
pdf_counter.configure(font=myFont)
pdf_counter.config(state="disabled")

send_btn = Button(window, text="ΑΠΟΣΤΟΛΗ", command=send_btn)
send_btn.place(rely=0.8, relx=0.5, x=0, y=0, anchor=CENTER)
send_btn.config(height=3, width=17)

exit_btn = Button(window, text="ΕΞΟΔΟΣ", command=window.destroy)
exit_btn.place(rely=0.9, relx=0.5, x=0, y=0, anchor=CENTER)

window.geometry("1280x600")
window.resizable(0, 0)

window.mainloop()