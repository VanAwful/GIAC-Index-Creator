# Requires the following be installed in python, if not already. 
# These can be installed using pip:
#
#   tkinter
#   pandas 
#   python-docx 
#   openpyxl
# 

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

def show_form():
    root = tk.Tk()
    root.title("Document Setup")
    root.geometry("450x400")
    result = {}

    def browse_file():
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
            title="Select the Input File"
        )
        if file_path:
            entry_input.delete(0, tk.END)
            entry_input.insert(0, file_path)
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".xlsx":
                var_format.set("Excel")
            elif ext == ".csv":
                var_format.set("CSV")

    tk.Label(root, text="Input File Path:").place(x=10, y=20)
    entry_input = tk.Entry(root, width=35)
    entry_input.place(x=120, y=20)
    tk.Button(root, text="Browse...", command=browse_file).place(x=380, y=18)

    tk.Label(root, text="Input Format:").place(x=10, y=60)
    var_format = tk.StringVar(value="Excel")
    tk.Radiobutton(root, text="Excel", variable=var_format, value="Excel").place(x=120, y=60)
    tk.Radiobutton(root, text="CSV", variable=var_format, value="CSV").place(x=200, y=60)

    tk.Label(root, text="Document Margins (cm):").place(x=10, y=100)
    tk.Label(root, text="Left:").place(x=10, y=130)
    tk.Label(root, text="Right:").place(x=10, y=160)
    tk.Label(root, text="Top:").place(x=10, y=190)
    tk.Label(root, text="Bottom:").place(x=10, y=220)

    entry_left = tk.Entry(root, width=10)
    entry_left.insert(0, "2.54")
    entry_left.place(x=120, y=130)
    entry_right = tk.Entry(root, width=10)
    entry_right.insert(0, "1.27")
    entry_right.place(x=120, y=160)
    entry_top = tk.Entry(root, width=10)
    entry_top.insert(0, "0.635")
    entry_top.place(x=120, y=190)
    entry_bottom = tk.Entry(root, width=10)
    entry_bottom.insert(0, "0.635")
    entry_bottom.place(x=120, y=220)

    def on_ok():
        result['InputFilePath'] = entry_input.get()
        result['InputFormat'] = var_format.get()
        result['Margins'] = {
            'Left': float(entry_left.get()),
            'Right': float(entry_right.get()),
            'Top': float(entry_top.get()),
            'Bottom': float(entry_bottom.get())
        }
        root.destroy()

    tk.Button(root, text="OK", command=on_ok, width=10).place(x=175, y=270)
    root.mainloop()
    return result if result else None

def get_output_file_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")],
        title="Save Output Word Document As",
        initialfile="IndexOutput.docx"
    )
    if not file_path:
        raise Exception("Output path not selected. Exiting script.")
    return file_path

def read_first_three_lines(input_file_path, input_format):
    if input_format == "CSV":
        with open(input_file_path, encoding="utf-8") as f:
            lines = [next(f).strip() for _ in range(3)]
    elif input_format == "Excel":
        df = pd.read_excel(input_file_path, header=None, nrows=3)
        lines = [", ".join(map(str, row)) for row in df.values]
    else:
        lines = []
    return lines

def get_if_headers(lines):
    prompt = "The first three lines of the document are:\n"
    prompt += "\n".join(lines)
    prompt += "\n\nDoes the first line contain headers?"
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno("Header Detection", prompt)
    return result

def format_data_by_first_column(data):
    return data.sort_values(by=data.columns[0], kind='stable')

def main():
    try:
        user_input = show_form()
        if not user_input:
            raise Exception("No input provided.")

        input_file_path = user_input['InputFilePath']
        input_format = user_input['InputFormat']
        margins = user_input['Margins']

        if not os.path.exists(input_file_path):
            raise Exception("Invalid input file path.")

        output_path = get_output_file_path()

        lines = read_first_three_lines(input_file_path, input_format)
        header_response = get_if_headers(lines)

        if input_format == "CSV":
            if header_response:
                data = pd.read_csv(input_file_path)
            else:
                data = pd.read_csv(input_file_path, header=None, names=["H1", "H2", "H3", "H4"])
        elif input_format == "Excel":
            if header_response:
                data = pd.read_excel(input_file_path, header=0)
            else:
                data = pd.read_excel(input_file_path, header=None, names=["H1", "H2", "H3", "H4"])
        else:
            raise Exception("Unsupported input format.")

        data = format_data_by_first_column(data)

        doc = Document()
        section = doc.sections[0]
        section.left_margin = Cm(margins['Left'])
        section.right_margin = Cm(margins['Right'])
        section.top_margin = Cm(margins['Top'])
        section.bottom_margin = Cm(margins['Bottom'])

        # Set two columns robustly
        cols = section._sectPr.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            section._sectPr.append(cols)
        cols.set(qn('w:num'), '2')

        previous_first_char = ''
        row_count = 0

        for _, row in data.iterrows():
            
            # Set a break point for faster testing. Remove or comment out for production use.
            if row_count >= 250:
                break

            topic = str(row.iloc[0]).lstrip()
            description = str(row.iloc[1]) if pd.notna(row.iloc[1]) else " "
            page = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            book = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
            first_char = topic[:1].upper() if topic else "#"
            if not first_char.isalpha():
                first_char = "#"

            if previous_first_char != first_char:
                if previous_first_char != '':
                    doc.add_page_break()
                p = doc.add_paragraph()
                run = p.add_run(first_char)
                run.bold = True
                run.font.size = Pt(24)
                run.font.name = 'Times New Roman'
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                previous_first_char = first_char

            bkpg = f" [bk {book}/pg{page}] "
            p = doc.add_paragraph()
            run1 = p.add_run(topic)
            run1.bold = True
            run1.font.size = Pt(10)
            run1.font.name = 'Times New Roman'

            run2 = p.add_run(bkpg)
            run2.italic = True
            run2.font.size = Pt(10)
            run2.font.name = 'Times New Roman'

            run3 = p.add_run(description)
            run3.font.size = Pt(10)
            run3.font.name = 'Times New Roman'

            p.space_after = Pt(8)
            row_count += 1

        doc.save(output_path)
        print(f"The Word document has been created successfully. Saved to {output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()