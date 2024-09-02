import zipfile
import os
from lxml import etree
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import json
import openai
import time

# Configura la chiave API di OpenAI
openai.api_key = "YOUR_API_KEY_HERE"

def extract_docx_xml(docx_path, extract_to):
    if not os.path.exists(docx_path):
        print("Errore: Inserisci un percorso valido per il file.")
        return False
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        docx_zip.extractall(extract_to)
        print(f"File estratti in: {extract_to}")
    return True

def read_xml_file(xml_path):
    with open(xml_path, 'rb') as file:
        xml_content = file.read()
    return xml_content

def parse_xml_content(xml_content):
    root = etree.fromstring(xml_content)
    return root

def get_chatgpt_response(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Sei un assistente che aiuta a modificare documenti Word."},
                {"role": "user", "content": prompt}
            ]
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        print(f"Errore durante la chiamata API: {e}")
        return None

def suggest_modifications(root):
    prompt = "Ecco gli elementi trovati nel documento Word:\n"
    for i, element in enumerate(root.iter(), 1):
        text = ''.join(element.itertext()).strip()
        if text:
            prompt += f"{i}. {text}\n"
    prompt += "\nGenera un elenco dei possibili campi che potrebbero essere modificabili nel documento Word fornito."
    return get_chatgpt_response(prompt)

def apply_changes(root, changes, checkboxes, xml_path):
    for i, element in enumerate(root.iter(), 1):
        if checkboxes.get(i, False):  # Se l'utente ha selezionato questa modifica
            new_text = changes.get(i)
            if new_text:
                for node in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    node.text = new_text
    save_xml_content(root, xml_path)

def save_xml_content(root, xml_path):
    with open(xml_path, 'wb') as file:
        file.write(etree.tostring(root, pretty_print=True))

def repackage_docx(xml_folder, new_docx_path):
    with zipfile.ZipFile(new_docx_path, 'w') as new_docx_zip:
        for foldername, subfolders, filenames in os.walk(xml_folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, xml_folder)
                new_docx_zip.write(file_path, arcname)
        print(f"Nuovo file docx creato: {new_docx_path}")

def load_file():
    filepath = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if not filepath:
        print("Errore: Inserisci un percorso valido per il file.")
        return
    extract_folder = 'extracted_content'
    if not extract_docx_xml(filepath, extract_folder):
        return

    document_xml_path = os.path.join(extract_folder, 'word', 'document.xml')
    document_xml_content = read_xml_file(document_xml_path)
    document_root = parse_xml_content(document_xml_content)

    suggestions = suggest_modifications(document_root)
    display_suggestions(suggestions, document_root, document_xml_path, extract_folder)

def display_suggestions(suggestions, root, xml_path, extract_folder):
    suggestion_text.delete(1.0, tk.END)
    suggestion_text.insert(tk.END, "Suggerimenti dall'AI per possibili modifiche:\n" + suggestions + "\n")

    suggestion_lines = suggestions.split("\n")
    row = 0
    for i, line in enumerate(suggestion_lines):
        if line.strip() and not line.startswith("Ecco una lista"):
            label = tk.Label(scroll_frame, text=line.strip(), anchor="w", justify="left")
            label.grid(row=row, column=0, sticky="w")
            entry = tk.Entry(scroll_frame, width=50)
            entry.grid(row=row, column=1, padx=10)
            checkbox_vars[i] = tk.BooleanVar()
            checkbox = tk.Checkbutton(scroll_frame, text="Modifica", variable=checkbox_vars[i])
            checkbox.grid(row=row, column=2)
            changes_entries[i] = entry
            row += 1

    apply_button.grid(row=row, column=0, columnspan=3)

def apply_changes_ui():
    changes = {}
    checkboxes = {}
    for i, entry in changes_entries.items():
        changes[i] = entry.get()
        checkboxes[i] = checkbox_vars[i].get()

    apply_changes(document_root, changes, checkboxes, document_xml_path)
    repackage_docx(extract_folder, 'Documento_Modificato.docx')
    print("Modifiche applicate e nuovo documento creato!")

root = tk.Tk()
root.title("Modifica Documento Word con AI")
root.geometry('800x600')

# Aggiunta di una barra di scorrimento orizzontale e verticale
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

file_button = tk.Button(frame, text="Carica File", command=load_file)
file_button.pack(anchor="nw", padx=10, pady=5)

suggestion_text = tk.Text(frame, height=10)
suggestion_text.pack(fill=tk.BOTH, expand=True)

scroll_canvas = tk.Canvas(frame)
scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scroll_frame = tk.Frame(scroll_canvas)
scroll_frame.bind("<Configure>", lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))

scrollbar = ttk.Scrollbar(frame, orient="vertical", command=scroll_canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill="y")
scroll_canvas.configure(yscrollcommand=scrollbar.set)
scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

changes_entries = {}
checkbox_vars = {}

apply_button = tk.Button(scroll_frame, text="Applica Modifiche", command=apply_changes_ui)
apply_button.grid(row=0, column=0, columnspan=3)

root.mainloop()
