from tkinter import *
from tkinter import filedialog, messagebox, ttk

from tkinter.filedialog import asksaveasfilename
from pypdf import PdfWriter, PdfReader
import pypdf


class PdfAttachmentApp:
    def __init__(self, master):
        """
        Initializes the PdfAttachmentApp class.

        Args:
        - master: the Tkinter master object.
        """
        self.master = master
        padding = 5
        master.title("Pdf Attachment")
        master.resizable(False, False)
        self.source_pdf_label = Label(master, text="Source PDF:", width=15)
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
        attachment_filename = filedialog.askopenfilenames(
            title="Select Attachment PDF", filetypes=[("PDF files", "*.pdf")]
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
            self.writer.append_pages_from_reader(PdfReader(self.source_pdf_filename))
            
            # bookmarks = (
            #     PdfReader(self.source_pdf_filename).outline
            #     if PdfReader(self.source_pdf_filename).outline
            #     else None
            # )
            # print(bookmarks)
            # self.writer._clone_outline(self.source_pdf_filename.outline)
            
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

            self.attachment_list = []
            self.source_pdf_filename = ""
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
