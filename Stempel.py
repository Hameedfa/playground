"""
Dieses Skript ermöglicht das schnelle Eintragen von Adressen, Datumsangaben und Titeln in PDF-Briefe.
Bitte beachten: Das PDF muss dafür über geeignete Formularfelder verfügen.
"""

from PyPDF2 import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def fill_pdf_form(template_path, output_path, replacements):
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)

    # Setze Formularwerte
    writer.update_page_form_field_values(writer.pages[0], replacements)

    # Speichern
    with open(output_path, "wb") as f_out:
        writer.write(f_out)

def run_gui():
    def generate_pdf():
        company = entry_company.get().strip()
        street = entry_street.get().strip()
        address = entry_address.get().strip()
        position = entry_position.get().strip()
        if not all([company, street, address, position]):
            messagebox.showerror("Fehler", "Bitte alle Felder ausfüllen.")
            return

        template_path = filedialog.askopenfilename(
            title="PDF-Vorlage wählen",
            filetypes=[("PDF-Dateien", "*.pdf")]
        )
        if not template_path:
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF-Dateien", "*.pdf")],
            title="Speichern unter"
        )
        if not output_path:
            return

        replacements = {
            "COMPANY_NAME": company,
            "STREET_ADDRESS": street,
            "ADDRESS": address,
            "POSITION": "Berwerbung als " + position, # Kann bei Bedarf entsprechend angepasst werden.
            "DATE": "Obertraubling, den " + datetime.now().strftime("%d.%m.%Y")
        }

        try:
            fill_pdf_form(template_path, output_path, replacements)
            messagebox.showinfo("Fertig", f"PDF erfolgreich gespeichert:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Irgendetwas ist schiefgelaufen:\n{e}")

    # GUI Setup
    root = tk.Tk()
    root.title("Stamp – Bewerbung automatisch ausfüllen")

    tk.Label(root, text="Firmenname:").grid(row=0, column=0, sticky="e")
    entry_company = tk.Entry(root, width=40)
    entry_company.grid(row=0, column=1)

    tk.Label(root, text="Straße + Nr:").grid(row=1, column=0, sticky="e")
    entry_street = tk.Entry(root, width=40)
    entry_street.grid(row=1, column=1)

    tk.Label(root, text="PLZ + Stadt:").grid(row=2, column=0, sticky="e")
    entry_address = tk.Entry(root, width=40)
    entry_address.grid(row=2, column=1)

    tk.Label(root, text="Stellenbezeichnung:").grid(row=3, column=0, sticky="e")
    entry_position = tk.Entry(root, width=40)
    entry_position.grid(row=3, column=1)

    tk.Button(root, text="PDF ausfüllen", command=generate_pdf).grid(row=4, column=0, columnspan=2, pady=10)

    root.mainloop()

run_gui()