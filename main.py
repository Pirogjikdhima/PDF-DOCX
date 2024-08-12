import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert
import os


def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("PDF files", "*.pdf"), ("DOCX files", "*.docx"), ("All files", "*.*"))
    )
    if file_path:
        file_path_var.set(file_path)


def convert_file():
    file_path = file_path_var.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a file.")
        return

    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == '.pdf':
        output_file = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=(("DOCX files", "*.docx"), ("All files", "*.*"))
        )
        if not output_file:
            return

        try:
            cv = Converter(file_path)
            cv.convert(output_file, start=0, end=None)
            cv.close()
            messagebox.showinfo("Success", f"PDF has been successfully converted to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    elif file_ext == '.docx':
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if not output_file:
            return

        try:
            convert(file_path, output_file)
            messagebox.showinfo("Success", f"DOCX has been successfully converted to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    else:
        messagebox.showerror("Error", "Unsupported file type. Please select a PDF or DOCX file.")



app = tk.Tk()
app.title("PDF <-> DOCX Converter")

# Set the window icon
icon_path = "Icon/icon.ico"
app.iconbitmap(icon_path)

file_path_var = tk.StringVar()

frame = tk.Frame(app)
frame.pack(pady=20, padx=20)

tk.Label(frame, text="Select file:").grid(row=0, column=0, padx=5, pady=5)
tk.Entry(frame, textvariable=file_path_var, width=40).grid(row=0, column=1, padx=5, pady=5)
tk.Button(frame, text="Browse", command=browse_file).grid(row=0, column=2, padx=5, pady=5)

tk.Button(app, text="Convert", command=convert_file).pack(pady=10)

app.mainloop()
