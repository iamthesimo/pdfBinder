"""
----------------------------- APP IDENTIFICATION --------------------------
    module:      pdfBinder
    file:        pdfBinder.py   
    date:        25 July 2023
    version:     1.0
    author:      Simone Santonoceto
------------------------------ APP DESCRIPTION ----------------------------
pdfBinder is the perfect tool for anyone who needs to attach files to PDF
"""

from tkinter import *
from tkinter import filedialog, messagebox, ttk

from tkinter.filedialog import asksaveasfilename
from pypdf import PdfWriter, PdfReader


def copy_bookmark(destination, bookmarks):
    """Copy bookmarks from source to destination"""
    """ Right now is only copying the bookmark but the link is not working"""
    for bookmark in bookmarks:
        if isinstance(bookmark, dict):
            destination.add_outline_item_dict(bookmark)
        else:
            copy_bookmark(destination, bookmark)


class PdfAttachmentApp:
    """Class implementation of SteINTE - BLE ATE application"""

    # Application constants
    __APPLICATION_VERSION__ = "1.0.0"
    __APPLICATION_NAME__ = "pdfBinder"
    __APPLICATION_ICON__ = "pdf-icon.ico"
    __ABOUT_IMAGE__ = "pdf-icon.ico"
    __ABOUT_SHORT_TXT__ = f"Version{__APPLICATION_VERSION__} {__APPLICATION_NAME__}"
    __DEVELOPER_NAME__ = "Simone Santonoceto"
    __DEVELOPER_CONTACTS__ = "simone.santonoceto@gmail.com"
    __ABOUT_DETAILED_TXT__ = (
        f"{__ABOUT_SHORT_TXT__}{__DEVELOPER_NAME__}{__DEVELOPER_CONTACTS__}"
    )
    __APPLICATION_COPYRIGHT__ = "GNU GPL v3.0"

    def __init__(self, master):
        """
        Initializes the PdfAttachmentApp class.

        Args:
        - master: the Tkinter master object.
        """
        self.master = master
        padding = 5
        master.title("pdfBinder")
        master.resizable(False, False)

        self.source_pdf_label = ttk.Label(master, text="Source PDF:", width=15)
        self.source_pdf_label.grid(row=0, column=0)

        self.source_pdf_entry = Entry(master, width=50)
        self.source_pdf_entry.grid(row=0, column=1)

        self.source_pdf_button = Button(
            master,
            text="Select file",
            command=self.select_source_pdf,
            width=15,
            padx=padding,
            pady=padding,
        )
        self.source_pdf_button.grid(row=0, column=2)

        self.attachment_label = Label(master, text="Attachment PDF:", width=15)
        self.attachment_label.grid(row=1, column=0)

        self.attachment_button = Button(
            master,
            text="Select file",
            command=self.select_attachment,
            width=15,
            padx=padding,
            pady=padding,
        )
        self.attachment_button.grid(row=1, column=2)

        self.attachment_listbox = Listbox(master, width=50)
        self.attachment_listbox.grid(row=1, column=1)
        self.attachment_listbox.bind("<Double-Button-1>", self.delete_attachment)
        self.attachment_listbox.bind("<Delete>", self.delete_attachment)
        self.attachment_listbox.bind("<BackSpace>", self.delete_attachment)
        self.attachment_listbox.bind("<MouseWheel>", self.mouse_wheel)

        self.generate_pdf_button = Button(
            master,
            text="Generate PDF",
            command=self.generate_pdf,
            width=40,
            padx=padding,
            pady=padding,
        )
        self.generate_pdf_button.grid(row=4, column=1)

        self.clear_attachments_button = Button(
            master,
            text="Clear Attachments",
            command=self.clear_attachments,
            width=15,
            padx=padding,
            pady=padding,
        )
        self.clear_attachments_button.grid(row=3, column=2)

        self.clear_button = Button(
            master,
            text="Clear All",
            command=self.clear_all,
            width=15,
            padx=padding,
            pady=padding,
        )
        self.clear_button.grid(row=4, column=2)

        self.attachment_list = []
        self.source_pdf_filename = ""
        self.writer = None

    def mouse_wheel(self, event):
        self.attachment_listbox.yview_scroll(-1, "units")

    def select_source_pdf(self):
        """
        Opens a file dialog to select the source PDF file.
        """
        self.source_pdf_filename = filedialog.askopenfilename(
            title="Select Source PDF", filetypes=[("PDF files", "*.pdf")]
        )
        if self.source_pdf_filename:
            self.source_pdf_entry.delete(0, END)
            self.source_pdf_entry.insert(0, self.source_pdf_filename.rsplit("/", 1)[1])

    def select_attachment(self):
        """
        Opens a file dialog to select an attachment PDF file.
        """
        filetypes = [("PDF files", "*.pdf"), ("All files", "*.*")]
        attachment_filename = filedialog.askopenfilenames(
            title="Select Attachment", filetypes=filetypes
        )
        try:
            if attachment_filename:
                for attachment in attachment_filename:
                    self.attachment_list.append(attachment)
                    self.attachment_listbox.insert(END, attachment.rsplit("/", 1)[1])
        except Exception as e:
            pass

    def add_attachment(self):
        """
        Adds the selected attachment PDF files to the writer object.
        """
        if self.writer:
            for attachment_filename in self.attachment_list:
                with open(attachment_filename, "rb") as pdf:
                    self.writer.add_attachment(attachment_filename, pdf.read())

    def clear_all(self):
        """Clears the source PDF file and attachment PDF files."""
        self.source_pdf_entry.delete(0, END)
        self.attachment_listbox.delete(0, END)

    def clear_attachments(self):
        """Clears the attachment PDF files."""
        self.attachment_listbox.delete(0, END)

    def delete_attachment(self, event):
        """Deletes the selected attachment PDF file."""
        self.attachment_listbox.delete(ANCHOR)

    def generate_pdf(self):
        """
        Generates a new PDF file with the selected source PDF file and attachment PDF files.
        """
        try:
            if not self.source_pdf_filename:
                messagebox.showerror("Error", "Please select a source PDF file")
                return
            self.writer = PdfWriter()
            self.reader = PdfReader(self.source_pdf_filename)

            self.writer.clone_document_from_reader(self.reader)

            self.add_attachment()
            source_pdf = self.source_pdf_filename.rsplit("/", 1)[1]
            source_folder = self.source_pdf_filename.rsplit("/", 1)[0]

            files = [("PDF files", "*.pdf")]
            destination_pdf = asksaveasfilename(
                filetypes=files,
                defaultextension=files,
                initialfile=source_pdf,
                initialdir=source_folder,
            )

            with open(destination_pdf, "wb") as f:
                self.writer.write(f)

            self.writer = None
            messagebox.showinfo(
                "PDF Generated",
                f"The PDF file {destination_pdf.rsplit('/', 1)[1]} has been generated successfully!",
            )
        except Exception as e:
            pass


root = Tk()
my_gui = PdfAttachmentApp(root)
root.mainloop()
