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

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tkinter.filedialog import asksaveasfilename
from pypdf import PdfWriter, PdfReader

class PdfAttachmentApp:
    """Class implementation of SteINTE - BLE ATE application"""

    # region Application constants
    __APPLICATION_VERSION__ = "1.0.0"
    __APPLICATION_NAME__ = "pdfBinder"
    __APPLICATION_ICON__ = "pdf-icon.ico"
    __ABOUT_IMAGE__ = "pdf-icon.ico"
    __ABOUT_SHORT_TXT__ = f"{__APPLICATION_NAME__} V.{__APPLICATION_VERSION__}"
    __DEVELOPER_NAME__ = "Simone Santonoceto"
    __DEVELOPER_CONTACTS__ = "simone.santonoceto@gmail.com"
    __ABOUT_DETAILED_TXT__ = f"{__ABOUT_SHORT_TXT__} {__DEVELOPER_NAME__} {__DEVELOPER_CONTACTS__}"
    __APPLICATION_COPYRIGHT__ = "GNU GPL v3.0"
    __APPLICATION_INFOS__ = "\n".join([__ABOUT_SHORT_TXT__,__DEVELOPER_NAME__,__DEVELOPER_CONTACTS__,__APPLICATION_COPYRIGHT__,])
    # endregion

    def __init__(self, root):
        """
        Initializes the PdfAttachmentApp class.

        Args:
        - master: the Tkinter master object.
        """
        self.root = root
        self.padding = 3
        self.root.title("pdfBinder")
        self.root.iconbitmap(self.__APPLICATION_ICON__)
        self.root.resizable(False, False)

        self.create_menu()
        self.create_main_area()        
    
    # GUI creation methods
    def create_menu(self, event=None):
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # File menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Select PDF", command=self.select_source_pdf, accelerator="Ctrl+S") # index 0
        self.file_menu.add_command(label="Select attachments", command=self.select_attachments, accelerator="Ctrl+A", ) # index 1
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Generate PDF", command=self.generate_pdf, accelerator="Enter", state=tk.DISABLED) # index 2
        self.file_menu.add_command(label="Clear attachments", command=self.clear_attachments,accelerator="Ctrl+D", state=tk.DISABLED) # index 3
        self.file_menu.add_command(label="Clear all", command=self.clear_all, accelerator="Ctrl+Shift+D", state=tk.DISABLED) # index 4
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.on_exit, accelerator="Esc") # index 5

        # Help menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        self.help_menu.add_command(label="Help", command=self.on_help, accelerator="F1")
        self.help_menu.add_command(label="About", command=self.on_about)

    def create_main_area(self):
        ''' Creates the top area of the GUI '''
        self.source_pdf_label = ttk.Label(root, text="Source PDF:", width=15)
        self.source_pdf_label.grid(row=0, column=0)

        self.source_pdf_entry = tk.Entry(root, width=50)
        self.source_pdf_entry.grid(row=0, column=1)

        self.source_pdf_button = tk.Button(
            root,
            text="Select file",
            command=self.select_source_pdf,
            width=15,
            padx=self.padding,
            pady=self.padding,
        )
        self.source_pdf_button.grid(row=0, column=2)

        self.attachment_label = tk.Label(root, text="Attachments:", width=15)
        self.attachment_label.grid(row=1, column=0)

        self.attachment_button = tk.Button(
            root,
            text="Select file",
            command=self.select_attachments,
            width=15,
            padx=self.padding,
            pady=self.padding,
        )
        self.attachment_button.grid(row=1, column=2)

        self.attachment_listbox = tk.Listbox(root, width=50, yscrollcommand=1)
        self.attachment_listbox.grid(row=1, column=1)

        self.generate_pdf_button = tk.Button(
            root,
            text="Generate PDF",
            command=self.generate_pdf,
            width=40,
            padx=self.padding,
            pady=self.padding,
            state=tk.DISABLED,
        )
        self.generate_pdf_button.grid(row=4, column=1)

        self.clear_attachments_button = tk.Button(
            root,
            text="Clear Attachments",
            command=self.clear_attachments,
            width=15,
            padx=self.padding,
            pady=self.padding,
            state=tk.DISABLED,
        )
        self.clear_attachments_button.grid(row=3, column=2)

        self.clear_all_button = tk.Button(
            root,
            text="Clear All",
            command=self.clear_all,
            width=15,
            padx=self.padding,
            pady=self.padding,
            state=tk.DISABLED,
        )
        self.clear_all_button.grid(row=4, column=2)

        self.attachment_list = []
        self.source_pdf_filename = ""
        self.writer = None
        
        # Key bindings
        self.root.bind("<F1>", self.on_help)
        self.root.bind("<Escape>", self.on_exit)
        self.root.bind("<Control_L><s>", self.select_source_pdf)
        self.root.bind("<Control_L><a>", self.select_attachments)

    # Event handlers
    def on_exit(self, event=None):
        self.root.destroy()

    def on_help(self, event=None):
        help_text = """This is a simple app created with Python and Tkinter to attach files to PDF:\n
  - Double click an attachment to delete it.
  - Click an attachment and press backspace to remove it.
  - Click an attachment and press delete to remove it.
  - Clear all to clear the source and attachments.
  - Clear attachments to clear the attachment."""
        messagebox.showinfo("Help", help_text,type="ok")

    def on_about(self):
        about_text = self.__APPLICATION_INFOS__
        messagebox.showinfo("About", about_text)

    # Buttons handlers
    def select_source_pdf(self, event=None):
        """
        Opens a file dialog to select the source PDF file.
        """
        self.source_pdf_filename = filedialog.askopenfilename(
            title="Select Source PDF", filetypes=[("PDF files", "*.pdf")]
        )
        if self.source_pdf_filename:
            self.source_pdf_entry.delete(0, tk.END)
            self.source_pdf_entry.insert(0, self.source_pdf_filename.rsplit("/", 1)[1])
        self.generate_pdf_button['state'] = tk.NORMAL
        self.file_menu.entryconfig("Generate PDF", state=tk.NORMAL)
        self.file_menu.entryconfig("Clear all", state=tk.NORMAL)
        self.clear_all_button['state'] = tk.NORMAL
        self.root.bind("<Return>", self.generate_pdf)
        self.root.bind("<Control_L><Shift_L><D>", self.clear_all)

    def select_attachments(self, event=None):
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
                    self.attachment_listbox.insert(tk.END, attachment.rsplit("/", 1)[1])
        
            self.clear_attachments_button['state'] = tk.NORMAL
            self.file_menu.entryconfig("Clear attachments", state=tk.NORMAL)
            self.root.bind("<Control_L><d>", self.clear_attachments)
            self.attachment_listbox.bind("<Double-Button-1>", self.delete_attachment)
            self.attachment_listbox.bind("<Delete>", self.delete_attachment)
            self.attachment_listbox.bind("<BackSpace>", self.delete_attachment)
            self.attachment_listbox.bind("<MouseWheel>", lambda event: self.attachment_listbox.yview_scroll(-1, "units"))
        except Exception as e:
            pass

    def insert_attachments(self):
        """
        Adds the selected attachment PDF files to the writer object.
        """
        if self.writer:
            for attachment_filename in self.attachment_list:
                with open(attachment_filename, "rb") as pdf:
                    self.writer.add_attachment(attachment_filename, pdf.read())

    def clear_all(self, event=None):
        """Clears the source PDF file and attachment PDF files."""
        if len(self.source_pdf_entry.get())>0:
            self.source_pdf_entry.delete(0, tk.END)
        self.clear_attachments()
        self.generate_pdf_button['state'] = tk.DISABLED
        self.clear_all_button['state'] = tk.DISABLED
        self.file_menu.entryconfig("Generate PDF", state=tk.DISABLED)
        self.file_menu.entryconfig("Clear all", state=tk.DISABLED)
        self.file_menu.entryconfig("Clear attachments", state=tk.DISABLED)
        self.root.unbind("<Return>")
        self.root.unbind("<Control_L><Shift_L><D>")

    def clear_attachments(self, event=None):
        """Clears the attachment PDF files."""
        if len(self.attachment_listbox.get(0, tk.END)):
            self.attachment_listbox.delete(0, tk.END)
        
        self.clear_attachments_button['state'] = tk.DISABLED
        self.file_menu.entryconfig("Clear attachments", state=tk.DISABLED)
        self.root.unbind("<Control_L><d>")
        self.attachment_listbox.unbind("<Double-Button-1>")
        self.attachment_listbox.unbind("<Delete>")
        self.attachment_listbox.unbind("<BackSpace>")
        self.attachment_listbox.unbind("<MouseWheel>")

    def delete_attachment(self, event):
        """Deletes the selected attachment PDF file."""
        self.attachment_listbox.delete(tk.ANCHOR)

    def generate_pdf(self, event=None):
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

            self.insert_attachments()
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
            print(Exception)

if __name__ == "__main__":
    root = tk.Tk()
    my_gui = PdfAttachmentApp(root)
    root.mainloop()
