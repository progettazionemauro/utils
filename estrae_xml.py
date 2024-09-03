import zipfile
import os
from lxml import etree
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox

def extract_docx_xml(docx_path, extract_to):
    """
    Extract the XML content from a .docx file (which is a zip file).
    """
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        docx_zip.extractall(extract_to)
        print(f"Files extracted to: {extract_to}")

def read_xml_file(xml_path):
    """
    Read the content of the specified XML file.
    """
    with open(xml_path, 'rb') as file:
        xml_content = file.read()
    return xml_content

def parse_xml_content(xml_content):
    """
    Parse the XML content.
    """
    root = etree.fromstring(xml_content)
    return root

def identify_highlighted_fields(root):
    """
    Identify fields that are highlighted in the document.
    """
    highlighted_fields = []
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    print("Debug: Starting to look for highlighted fields...")

    for elem in root.findall('.//w:r', namespaces=namespace):
        shading = elem.find('.//w:rPr/w:shd', namespaces=namespace)

        if shading is not None:
            fill_val = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
            print(f"Debug: Found a shading with fill value '{fill_val}'")

            if fill_val and fill_val not in ['auto', 'clear']:
                text_elem = elem.find('.//w:t', namespaces=namespace)
                if text_elem is not None:
                    text = ''.join(text_elem.itertext()).strip()
                    if text and text not in highlighted_fields:
                        print(f"Debug: Adding highlighted field: {text}")
                        highlighted_fields.append(text)

    return highlighted_fields

def update_xml_content(root, changes):
    """
    Update the XML content with the new changes.
    """
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for elem in root.findall('.//w:r', namespaces=namespace):
        shading = elem.find('.//w:rPr/w:shd', namespaces=namespace)

        if shading is not None:
            fill_val = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')

            if fill_val and fill_val not in ['auto', 'clear']:
                text_elem = elem.find('.//w:t', namespaces=namespace)
                if text_elem is not None:
                    original_text = ''.join(text_elem.itertext()).strip()
                    if original_text in changes:
                        print(f"Debug: Replacing '{original_text}' with '{changes[original_text]}'")
                        text_elem.text = changes[original_text]

def save_xml_content(root, xml_path):
    """
    Save the modified XML content back to the file.
    """
    with open(xml_path, 'wb') as file:
        file.write(etree.tostring(root, pretty_print=True))

def repackage_docx(xml_folder, new_docx_path):
    """
    Repackage the extracted XML content back into a .docx file.
    """
    with zipfile.ZipFile(new_docx_path, 'w') as new_docx_zip:
        for foldername, subfolders, filenames in os.walk(xml_folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, xml_folder)
                new_docx_zip.write(file_path, arcname)
        print(f"New modified file created: {new_docx_path}")

def display_suggestions(suggestions, root, xml_path, folder):
    """
    Display suggestions in the Tkinter interface with input fields for changes.
    """
    for widget in frame.winfo_children():
        widget.destroy()  # Clear previous widgets

    tk.Label(frame, text="Suggestions from AI for possible modifications:").pack(pady=5)

    entry_widgets.clear()  # Clear previous entries

    if suggestions:
        for suggestion in suggestions:
            label_frame = tk.Frame(frame)
            label_frame.pack(fill=tk.X, pady=2)

            tk.Label(label_frame, text=suggestion).pack(side=tk.LEFT, padx=5)
            entry = tk.Entry(label_frame, width=50)
            entry.pack(side=tk.RIGHT, padx=5)

            entry_widgets.append((suggestion, entry))
    else:
        tk.Label(frame, text="No highlighted text found in the document.").pack(pady=10)

    # Apply changes button
    apply_button = ttk.Button(frame, text="Applica Modifiche", command=lambda: apply_changes(root, xml_path, folder))
    apply_button.pack(pady=10)

def apply_changes(xml_root, xml_path, folder):
    """
    Apply the changes from the input fields to the XML content and save the modified document.
    """
    changes = {}
    for suggestion, entry in entry_widgets:
        new_value = entry.get()
        if new_value.strip():
            changes[suggestion] = new_value.strip()

    # Update XML content with new changes
    update_xml_content(xml_root, changes)
    
    # Save the modified XML back to the file
    save_xml_content(xml_root, xml_path)
    
    # Repackage the modified XML back into a .docx file
    new_docx_path = os.path.join(folder, "Modified_Document.docx")
    repackage_docx(folder, new_docx_path)

    messagebox.showinfo("Success", "Modifications applied and new document saved!")

    root.quit()  # Corrected line to close the Tkinter window

def load_file():
    """
    Load the Word file using a file dialog.
    """
    global document_root, document_xml_path, extract_folder  # Make these variables global

    filepath = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    
    if filepath:
        extract_folder = 'extracted_content'
        extract_docx_xml(filepath, extract_folder)
        
        document_xml_path = os.path.join(extract_folder, 'word', 'document.xml')
        document_xml_content = read_xml_file(document_xml_path)
        document_root = parse_xml_content(document_xml_content)
        
        suggestions = identify_highlighted_fields(document_root)
        display_suggestions(suggestions, document_root, document_xml_path, extract_folder)

# Tkinter GUI Setup
root = tk.Tk()
root.title("Modify Word Document with AI")
root.geometry("600x400")  # Set a larger initial size for the window

frame = ttk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

file_button = ttk.Button(frame, text="Carica File", command=load_file)
file_button.pack(pady=5)

entry_widgets = []  # List to store suggestion and entry widget pairs

root.protocol("WM_DELETE_WINDOW", root.quit)  # Correctly handle the window close button

root.mainloop()
