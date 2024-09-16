import sys
import re
import win32com.client
import tkinter as tk
from tkinter import messagebox
import os
import comtypes.client
from PyPDF2 import PdfMerger
import shutil

def main():
    # if len(sys.argv) < 2:
    #     print("Usage: python script.py <emailID>")
    #     sys.exit(1)

#     emailID = sys.argv[1]
    
    ## HADR CODE for testing
    emailID = "00000000E08501EAD5D9744B869F96120D489BA40700392886B1971E2E4C8C4CC3A714A1DB7700D5012603B600006CD8F4536658204BA47FCD18199852B8000007FAC9110000"

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the email item using EntryID
    mail_item = outlook.GetItemFromID(emailID)

    # Read the email body
    body = mail_item.Body

    # Search for regex pattern
    match = re.search(r'\b\d+/\d+/\d{4}\b', body)
    if match:
        found_number = match.group()
    else:
        found_number = "Enter File Name"

    # Create Tkinter window
    root = tk.Tk()
    root.title("Save PDF")
    root.geometry("400x100")

    tk.Label(root, text="Save Path:").pack(pady=5)
    entry = tk.Entry(root, width=50)
    entry.insert(0, found_number)
    entry.pack(pady=5)

    def on_save():
        save_name = entry.get()
        if save_name:
            save_name = save_name.replace("/", "-")
            create_pdf(mail_item, save_name)
            messagebox.showinfo("Success", f"PDF created: {save_name}.pdf")
            root.destroy()
        else:
            messagebox.showwarning("Input Error", "Please enter a valid file name.")

    tk.Button(root, text="Save", command=on_save).pack(pady=5)

    root.mainloop()

def create_pdf(mail_item, save_name):
    # Create a folder to store temporary files
    temp_folder = os.path.join(os.getcwd(), "temp_pdf_files")
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)

    pdf_files = []

    # Prepare email information
    sender = mail_item.SenderEmailAddress
    recipients = ', '.join([rec.Address for rec in mail_item.Recipients if rec.Type == 1])  # Type 1 = To
    cc_recipients = ', '.join([rec.Address for rec in mail_item.Recipients if rec.Type == 2])  # Type 2 = CC
    subject = mail_item.Subject
    sent_on = mail_item.SentOn.strftime("%d/%m/%Y %H:%M:%S")  # Adjust date format as needed

    # Construct HTML content with additional information
    email_html_content = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .email-header {{ margin-bottom: 20px; }}
            .email-header h2 {{ margin: 0; }}
            .email-header p {{ margin: 5px 0; }}
        </style>
    </head>
    <body>
        <div class="email-header">
            <h2>Email Details</h2>
            <p><strong>Trimis:</strong> {sent_on}</p>
            <p><strong>De la:</strong> {sender}</p>
            <p><strong>Către:</strong> {recipients}</p>
            <p><strong>CC:</strong> {cc_recipients}</p>
            <p><strong>Subiect:</strong> {subject}</p>
        </div>
        <hr>
        {mail_item.HTMLBody}
    </body>
    </html>
    """

    # Save email body as HTML
    email_html_path = os.path.join(temp_folder, "email.html")
    with open(email_html_path, 'w', encoding='utf-8') as f:
        f.write(email_html_content)

    # Convert HTML to PDF using Word
    email_pdf_path = os.path.join(temp_folder, "email.pdf")
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(email_html_path)

    # Handle potential encoding issues
    doc.Activate()
    word.Selection.WholeStory()
    word.Selection.LanguageID = 1033  # Set language to English to avoid encoding issues

    # wdExportFormatPDF = 17
    doc.SaveAs(email_pdf_path, FileFormat=17)
    doc.Close(False)
    word.Quit()
    del doc
    del word

    pdf_files.append(email_pdf_path)

    # Process attachments
    for attachment in mail_item.Attachments:
        filename = attachment.FileName
        if filename.lower().endswith(('.pdf', '.docx')):
            attachment_path = os.path.join(temp_folder, filename)
            attachment.SaveAsFile(attachment_path)

            if filename.lower().endswith('.docx'):
                # Convert DOCX to PDF using Word
                word = comtypes.client.CreateObject('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(attachment_path)
                pdf_attachment_path = attachment_path[:-5] + '.pdf'

                # Handle potential encoding issues
                doc.Activate()
                word.Selection.WholeStory()
                word.Selection.LanguageID = 1033  # Set language to English

                doc.SaveAs(pdf_attachment_path, FileFormat=17)
                doc.Close(False)
                word.Quit()
                del doc
                del word
                pdf_files.append(pdf_attachment_path)
            else:
                pdf_files.append(attachment_path)

    # Merge PDFs
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)

    output_pdf_path = os.path.join(os.getcwd(), f"{save_name}.pdf")
    merger.write(output_pdf_path)
    merger.close()

    # Cleanup
    shutil.rmtree(temp_folder)

if __name__ == "__main__":
    main()
