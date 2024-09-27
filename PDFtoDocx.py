#Welcome to the PDF to Docx Converting Tool code!
#Before using this code please remember to install
#*pip install pywin32
#*pip install tkinterdnd2

#The editing tool will show a drag box in which to drop your PDF files
#And browse files to output the converted Docx documents.

import os
import win32com.client
from tkinter import Tk, Frame, Label, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

def convert_pdf_to_docx(pdf_path, output_dir):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        input_file = os.path.abspath(pdf_path)
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File not found: {input_file}")

        wb = word.Documents.Open(input_file)
        output_file = os.path.join(output_dir, os.path.basename(pdf_path).replace('.pdf', '.docx'))
        wb.SaveAs2(output_file, FileFormat=16)
        wb.Close()
        word.Quit()

        print(f"Successfully converted {pdf_path} to {output_file}")
        return output_file

    except Exception as e:
        print(f"Error converting {pdf_path}: {e}")
        return None


def on_drop(event):
    # Get the files dropped into the window
    file_paths = event.data.split('}')  # Split based on closing brace for multiple files
    file_paths = [file.strip('{} ') for file in file_paths if file]  # Clean up the file paths
    print(f"Files dropped: {file_paths}")

    # Ask for the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    if not output_dir:
        print("No output directory selected.")
        return

    # Convert each PDF to DOCX
    for file in file_paths:
        if file.endswith('.pdf'):
            output_file = convert_pdf_to_docx(file, output_dir)
            if output_file:
                print(f"Converted file saved at: {output_file}")
            else:
                print(f"Failed to convert file: {file}")
        else:
            print(f"Unsupported file type: {file}")

    messagebox.showinfo("Conversion Complete", "All files have been converted.")


def create_gui():
    # Create main window
    root = TkinterDnD.Tk()
    root.title("PDF to DOCX Converter")

    # Frame for the drag-and-drop area
    frame = Frame(root, width=400, height=200)
    frame.pack(padx=10, pady=10)

    # Label for instructions
    label = Label(frame, text="Drag and drop PDF files here", padx=10, pady=10)
    label.pack(pady=10)

    # Register the drop target
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

    # Start the Tkinter main loop
    print("Starting the GUI...")
    root.mainloop()

if __name__ == "__main__":
    create_gui()