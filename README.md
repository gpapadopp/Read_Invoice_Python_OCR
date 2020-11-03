# Read_Invoice_and_Send_Email

ENGLISH VERISON

This application was created to send invoices and facilitate the user in managing his customer base, as well as the invoices of his customers.

The possibilities of the application are the following:

1.	Import a client list in Excel format (.xlsx) for reading the names, VAT numbers and e-mails of customers.
2.	Import a PDF file (.pdf) to read invoices. Invoices may have been scanned by a scanner, or exported from an invoice generator. (The program is optimized for importing a scanned invoice.)

Program operation:

1.	The user must first import the above files. The client list, in Excel format (.xlsx), and the invoices, in PDF format (.pdf). In the corresponding fields of the program.
2.	Then the user must press the "Send" button to start the process of executing the algorithm.
3.	After pressing the "Send" button, the program asks the user to enter the subject and the main text of the e-mail that will be sent.
4.	After typing the above, then the program will execute its algorithm, where it works as follows:
  a.	The program will create a folder on the user's local C: disk named "temp_pdf_page_by_page".
  b.	It will then read the invoice PDF file, and divide it according to its pages into other PDFs called "document-pagei.pdf", where i is the page number of the original PDF. That is, for each page of the PDF file, it will create a one-page PDF file in the above folder, with the above name.
  c.	After splitting the original PDF file into several one-page PDF files, it will create JPG files for each page, for each one-page PDF file.
  d.	These photos will be fed back to the OCR system built into the program. Which will create a list for each page of the PDF file, with the contents of the page, i.e. the "objects" that the OCR of the program has decoded.
  e.	Having created a list that contains as many lists as the pages of our original PDF, then the program makes an effort to "open" the Excel file (.xlsx).
  f.	So, after the program has loaded the Excel file (.xlsx), then it reads and saves in a list, all the VAT numbers from the Q list of Excel (.xlsx).
  g.	The program then performs multiple checks to see if any of the VAT numbers read in the step above correspond to any of the contents of the list exported by the OCR program.
    i.	If there is one of the VATs of the list in step f. in the list exported by the OCR program, then creates and adds data to 3 parallel lists. The contents of these lists will be as follows. The first list contains the emails of the customers. The second list will contain the page number that the program should send. The third list will contain the VAT numbers.
    ii.	If one of the VATs of the list in step f does not equals in any element of the list exported by the OCR program, then the program does not add data to any of the 3 parallel lists, and proceeds to the next comparison. In case there are no other comparisons to make, i.e., he has accessed all the elements of the list of VATs of step f. and none is equals with any element of the list exported by the OCR program, then the 3 parallel lists remain empty.
  h.	Once the parallel data lists have been created the program then performs a repetition with the number of repetitions as well as the number of lists that the clients' emails have been registered. In the repetition, the program calls a function, called "send_email", in which our data will pass, in the following order, e-mail client, e-mail subject, mainly e-mail text, location of the attached file. The client's e-mail is in the 1st parallel list of the program (see step g.). The subject and the main text of the e-mail, where the user of the program has entered it in the initial steps. The location of the attached file is created from the default path position of the PDF file, (see step b.) where in "i", the program will put the number in the 2nd parallel list (see step g.) reduced by one unit.
  i.	In the "send_email" function, you create the direct connection to Microsoft SMTP Server, because the sender's e-mail is based on e-mail with a "Hotmail" extension. Once the connection with the SMTP Server is achieved with the appropriate Credentials of the sender, then the data that should be sent to the client is sent, i.e., the subject of the e-mail, the main text of the e-mail, and as an attachment, the appropriate invoice page corresponding to each customer.
      In case it fails to do any of the above, sending the email to the customer is considered a failure, so it will create and add to an independent list, all the paths of the attachments that the function failed to send.
  j.	After the repetitions of step h are completed, then the program will look for contents in the list created by the "send_email" function in case no e-mail is sent.
    i.	If there is content in this list, then the program will create a folder on the user's local C: disk named "INVOICES_NOT_SEND". In this folder, it will create a text file (.txt) whereas contents it will have the contents of the list created by the "send_email" function in case no e-mail is sent, that is, the locations of the attachments that failed the program to send.
    ii.	If there is no content in this list, then the program will not create any folder, and its operation will be completed.
  k.	The program after performing all the above steps, will then return a success message to the user, where with the help of the graphical user interface, the above message will appear in a "message - information box".

Programming language: Python v.3.8.5
Programming environment: PyCharm v.2020.2


GREEK VERSION

Η παρόν εφαρμογή δημιουργήθηκε για την αποστολή τιμολογίων και την διευκόλυνση του χρήστη στην διαχείριση του πελατολογίου του, καθώς και των τιμολογίων των πελατών του.

Οι δυνατότητες της εφαρμογής είναι οι εξής:

1.	Εισαγωγή ενός πελατολογίου σε μορφή Excel (.xlsx) για την ανάγνωση των ονομάτων, των ΑΦΜ και των e-mail των πελατών.
2.	Εισαγωγή ενός αρχείου PDF (.pdf) για την ανάγνωση των τιμολογίων. Τα τιμολόγια μπορεί να έχουν σαρωθεί με κάποιο scanner, ή να έχουν γίνει εξαγωγή από κάποιο πρόγραμμα δημιουργίας τιμολογίων. (Το πρόγραμμα έχει βελτιστοποιηθεί για την εισαγωγή ενός σαρωμένου τιμολογίου.)

Λειτουργία του προγράμματος:

1.	Ο χρήστης καταρχάς, θα πρέπει να εισάγει τα παραπάνω αρχεία. Το πελατολόγιο, σε μορφή Excel (.xlsx), και τα τιμολόγια, σε μορφή PDF (.pdf). Στα αντίστοιχα πεδία του προγράμματος.
2.	Έπειτα ο χρήστης θα πρέπει να πατήσει το κουμπί «Αποστολή», ώστε να ξεκινήσει η διαδικασία εκτέλεσης του αλγορίθμου.
3.	Μετά το πάτημα του κουμπιού «Αποστολή», το πρόγραμμα ζητάει από τον χρήστη να πληκτρολογήσει το θέμα και το κυρίως κείμενο του e-mail που θα αποσταλεί. 
4.	Αφού πληκτρολογήσει τα παραπάνω, τότε το πρόγραμμα θα εκτελέσει τον αλγόριθμο του, όπου λειτουργεί ως εξής:
  a.	Το πρόγραμμα θα δημιουργήσει έναν φάκελο στον τοπικό δίσκο C: του συστήματος του χρήστη, με όνομα «temp_pdf_page_by_page».
  b.	Έπειτα, θα διαβάσει το αρχείο PDF με τα τιμολόγια, και θα χωρίσει ανάλογα με τις σελίδες του, σε άλλα PDF με το όνομα «document-pagei.pdf», όπου i, ο αριθμός της σελίδας του αρχικού PDF. Δηλαδή για κάθε μια σελίδα του αρχείου PDF, θα δημιουργήσει ένα μονοσέλιδο αρχείο PDF μέσα στο παραπάνω φάκελο, με το παραπάνω όνομα.
  c.	Αφού έχει χωρίσει το αρχικό αρχείο PDF σε πολλά μονοσέλιδα αρχεία PDF, τότε για κάθε μια σελίδα, για κάθε ένα μονοσέλιδο αρχείο PDF, θα δημιουργήσει αρχεία JPG, δηλαδή φωτογραφίες από το κάθε PDF.
  d.	Αυτές τις φωτογραφίες θα τις ανατροφοδοτήσει στο σύστημα OCR που έχει ενσωματωθεί στο πρόγραμμα. Το οποίο θα δημιουργήσει μια λίστα για κάθε σελίδα του αρχείου PDF, με τα περιεχόμενα της σελίδας, δηλαδή τα «αντικείμενα» που έχει αποκωδικοποιήσει το OCR του προγράμματος.
  e.	Αφού έχει δημιουργήσει μια λίστα η οποία περιέχει τόσες λίστες όσες οι σελίδες του αρχικού PDF μας, τότε το πρόγραμμα κάνει μια προσπάθεια να «ανοίξει» το αρχείο Excel (.xlsx).
  f.	Αφού λοιπόν το πρόγραμμα έχει φορτώσει το αρχείο Excel (.xlsx), τότε διαβάζει και αποθηκεύει σε μια λίστα, όλα τα ΑΦΜ από την λίστα Q του Excel (.xlsx).
  g.	Έπειτα, το πρόγραμμα κάνει πολλαπλούς ελέγχους με στόχο να βρει αν κάποιο από τα ΑΦΜ που έχει διαβάσει στο ακριβώς παραπάνω βήμα, αντιστοιχεί σε κάποιο από τα περιεχόμενα της λίστας που έχει εξάγει το OCR του προγράμματος.
    i.	Αν υπάρχει κάποιο από τα ΑΦΜ της λίστας του βήματος f. στην λίστα που έχει εξάγει το OCR του προγράμματος, τότε δημιουργεί και προσθέτει δεδομένα σε 3 παράλληλες λίστες. Τα περιεχόμενα αυτών των λιστών θα είναι ως εξής. Η πρώτη λίστα περιέχει τα e-mail των πελατών. Η δεύτερη λίστα θα περιέχει τον αριθμό της σελίδας που θα πρέπει να στείλει το πρόγραμμα. Η τρίτη λίστα, θα περιέχει τα ΑΦΜ.
    ii.	Αν δεν υπάρχει κάποιο από τα ΑΦΜ της λίστας του βήματος f. στην λίστα που έχει εξάγει το OCR, τότε το πρόγραμμα δεν προσθέτει δεδομένα σε καμία από τις 3 παράλληλες λίστες, και προχωράει στην επόμενη σύγκριση. Στην περίπτωση που δεν υπάρχουν άλλες συγκρίσεις να κάνει, δηλαδή έχει προσπελάσει όλα τα στοιχεία της λίστας των ΑΦΜ του βήματος f., τότε οι 3 παράλληλες λίστες μένουν κενές.
  h.	Αφού έχουν δημιουργηθεί οι παράλληλες λίστες με τα δεδομένα σε μια σειρά, τότε το πρόγραμμα εκτελεί μια επανάληψη με αριθμό επαναλήψεων όσο και το πλήθος της λίστας που έχει καταχωρημένα τα e-mail των πελατών. Στην επανάληψη, το πρόγραμμα καλεί μια συνάρτηση, με όνομα «send_email», στην οποία θα περάσει τα δεδομένα μας, με την εξής σειρά, e-mail πελάτη, θέμα του e-mail, κυρίως κείμενο του e-mail, τοποθεσία του επισυναπτόμενου αρχείου. Το e-mail του πελάτη, υπάρχει στην 1η παράλληλη λίστα του προγράμματος (βλέπε βήμα g.). Το θέμα και το κυρίως κείμενο του e-mail, όπου το έχει πληκτρολογήσει το χρήστης του προγράμματος στα αρχικά βήματα. Η τοποθεσία του επισυναπτόμενου αρχείου, δημιουργείται από την προκαθορισμένη θέση της σελίδας του αρχείου PDF, (βλέπε βήμα b.) όπου στο «i», το πρόγραμμα θα βάλει τον αριθμό που βρίσκεται στην 2η παράλληλη λίστα (βλέπε βήμα g.) μειωμένο κατά μία μονάδα.
  i.	Στην συνάρτηση «send_email», δημιουργείτε η απευθείας σύνδεση με τον SMTP Server της Microsoft, διότι το e-mail του αποστολέα είναι βασισμένο σε e-mail με κατάληξη «Hotmail». Αφού επιτευχθεί η σύνδεση με τον SMTP Server με τα κατάλληλα Credentials του αποστολέα, τότε γίνεται αποστολή των δεδομένων που θα πρέπει να αποσταλούν στον πελάτη, αυτά είναι, το θέμα του e-mail, το κυρίως κείμενο του e-mail, και σαν επισυναπτόμενο, η κατάλληλη σελίδα του τιμολογίου που αντιστοιχεί σε κάθε πελάτη.
      Στην περίπτωση που δεν καταφέρει να κάνει κάτι από τα παραπάνω, η αποστολή του e-mail στον πελάτη θεωρείται αποτυχημένη, και άρα θα δημιουργήσει και θα προσθέσει σε μια ανεξάρτητη λίστα, όλες τις τοποθεσίες των επισυναπτόμενων που δεν κατάφερε να η συνάρτηση να στείλει.
  j.	Αφότου ολοκληρωθούν οι επαναλήψεις του βήματος h., τότε το πρόγραμμα θα αναζητήσει αν υπάρχουν περιεχόμενα στην λίστα που δημιουργεί η συνάρτηση «send_email» στην περίπτωση που δεν αποσταλεί κάποιο e-mail.
    i.	Αν υπάρχουν περιεχόμενα στην λίστα αυτή, τότε το πρόγραμμα θα δημιουργήσει έναν φάκελο στον τοπικό δίσκο C: του συστήματος του χρήστη, με όνομα «INVOICES_NOT_SEND». Μέσα σε αυτόν τον φάκελο, θα δημιουργήσει ένα αρχείο κειμένου (.txt) όπου σαν περιεχόμενα θα έχει τα περιεχόμενα της λίστα που δημιουργεί η συνάρτηση «send_email» στην περίπτωση που δεν αποσταλεί κάποιο e-mail, δηλαδή, τις τοποθεσίες των επισυναπτόμενων που δεν κατάφερε το πρόγραμμα να στείλει.
    ii.	Αν δεν υπάρχουν περιεχόμενα στην λίστα αυτή, τότε το πρόγραμμα δεν θα δημιουργήσει κανέναν φάκελο, και θα ολοκληρωθεί η λειτουργία του.
  k.	Το πρόγραμμα αφότου εκτελέσει όλα τα παραπάνω βήματα, τότε θα επιστρέψει ένα μήνυμα επιτυχία στον χρήστη, όπου με την βοήθεια της γραφικής διεπαφής του χρήστη, θα εμφανιστεί το παραπάνω μήνυμα, σε ένα «πλαίσιο μηνυμάτων – πληροφοριών».

Γλώσσα προγραμματισμού : Python v.3.8.5
Προγραμματιστικό Περιβάλλον : PyCharm v.2020.2

