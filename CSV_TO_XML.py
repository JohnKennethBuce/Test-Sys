import csv
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox


def csv_to_xml(csv_file_path, xml_file_path):
    try:
        root = ET.Element("Records")

        with open(csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)

            for row in reader:
                item = ET.SubElement(root, "Record")
                for key, value in row.items():
                    child = ET.SubElement(item, key)
                    child.text = value

        tree = ET.ElementTree(root)
        tree.write(xml_file_path, encoding='utf-8', xml_declaration=True)

        messagebox.showinfo("Success", f"XML file saved to:\n{xml_file_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def select_and_convert():
    csv_file = filedialog.askopenfilename(
        title="Select CSV File", filetypes=[("CSV Files", "*.csv")]
    )
    if not csv_file:
        return

    xml_file = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")],
        title="Save XML File"
    )
    if not xml_file:
        return

    csv_to_xml(csv_file, xml_file)


# ===== GUI Setup =====
window = tk.Tk()
window.title("CSV to XML Converter")
window.geometry("400x200")
window.resizable(False, False)

label = tk.Label(window, text="CSV to XML Converter", font=("Arial", 16))
label.pack(pady=20)

button = tk.Button(window, text="Select CSV and Convert", command=select_and_convert, width=25, height=2)
button.pack(pady=10)

exit_button = tk.Button(window, text="Exit", command=window.destroy, width=10)
exit_button.pack(pady=10)

window.mainloop()
