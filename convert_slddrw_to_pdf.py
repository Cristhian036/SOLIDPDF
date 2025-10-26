

import os
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger
import tempfile
import threading

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertir SLDDRW a PDF")
        self.archivos = []

        self.frame = tk.Frame(root, padx=10, pady=10)
        self.frame.pack()

        self.btn_seleccionar = tk.Button(self.frame, text="Seleccionar archivos", command=self.seleccionar_archivos)
        self.btn_seleccionar.grid(row=0, column=0, pady=5, sticky="ew")

        self.btn_convertir = tk.Button(self.frame, text="Convertir a PDF", command=self.iniciar_conversion, state="disabled")
        self.btn_convertir.grid(row=1, column=0, pady=5, sticky="ew")

        self.progress = ttk.Progressbar(self.frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=2, column=0, pady=10)

        self.status = tk.Label(self.frame, text="Seleccione archivos para comenzar.")
        self.status.grid(row=3, column=0, pady=5)

    def seleccionar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Selecciona los planos .SLDDRW",
            filetypes=[("SolidWorks Drawings", "*.slddrw")]
        )
        self.archivos = list(archivos)
        if self.archivos:
            self.status.config(text=f"{len(self.archivos)} archivo(s) seleccionado(s).")
            self.btn_convertir.config(state="normal")
        else:
            self.status.config(text="No se seleccionaron archivos.")
            self.btn_convertir.config(state="disabled")

    def iniciar_conversion(self):
        self.btn_convertir.config(state="disabled")
        self.btn_seleccionar.config(state="disabled")
        self.progress['value'] = 0
        self.status.config(text="Iniciando conversión...")
        threading.Thread(target=self.convertir_archivos).start()

    def convertir_archivos(self):
        if not self.archivos:
            self.status.config(text="No se seleccionaron archivos.")
            self.btn_convertir.config(state="disabled")
            self.btn_seleccionar.config(state="normal")
            return
        try:
            swApp = win32com.client.Dispatch('SldWorks.Application')
            swApp.Visible = False
        except Exception as e:
            self.status.config(text=f"Error iniciando SolidWorks: {e}")
            self.btn_seleccionar.config(state="normal")
            return
        with tempfile.TemporaryDirectory() as temp_dir:
            pdfs = []
            total = len(self.archivos)
            for idx, slddrw_path in enumerate(self.archivos, 1):
                pdf = exportar_a_pdf(swApp, slddrw_path, temp_dir)
                if pdf:
                    pdfs.append(pdf)
                self.progress['value'] = (idx / total) * 100
                self.status.config(text=f"Procesando {idx}/{total}: {os.path.basename(slddrw_path)}")
                self.root.update_idletasks()
            swApp.ExitApp()
            if not pdfs:
                self.status.config(text="No se generaron PDFs.")
                self.btn_seleccionar.config(state="normal")
                return
            output_folder = os.path.dirname(self.archivos[0])
            output_pdf = os.path.join(output_folder, "Planos_Combinados.pdf")
            merger = PdfMerger()
            for pdf in pdfs:
                merger.append(pdf)
            merger.write(output_pdf)
            merger.close()
            self.status.config(text=f"PDF combinado guardado en: {output_pdf}")
            messagebox.showinfo("Éxito", f"PDF combinado guardado en:\n{output_pdf}")
        self.btn_seleccionar.config(state="normal")
        self.btn_convertir.config(state="normal")

def exportar_a_pdf(swApp, slddrw_path, temp_dir):
    from win32com.client import VARIANT
    import pythoncom
    swDocDRAWING = 3
    swOpenDocOptions_Silent = 64
    errors = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    drawing = swApp.OpenDoc6(slddrw_path, swDocDRAWING, swOpenDocOptions_Silent, '', errors, warnings)
    if drawing is None:
        print(f"No se pudo abrir {os.path.basename(slddrw_path)}")
        return None
    pdf_path = os.path.join(temp_dir, os.path.splitext(os.path.basename(slddrw_path))[0] + '.pdf')
    try:
        result = drawing.SaveAs(pdf_path)
        if result:
            print(f"Guardado: {pdf_path}")
            return pdf_path
        else:
            print(f"Error al guardar {pdf_path}")
            return None
    except Exception as e:
        print(f"Error exportando {os.path.basename(slddrw_path)}: {e}")
        return None
    finally:
        swApp.CloseDoc(os.path.basename(slddrw_path))


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
