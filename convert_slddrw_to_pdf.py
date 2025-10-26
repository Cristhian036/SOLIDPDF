
import os
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfMerger
import tempfile

def seleccionar_archivos():
    root = tk.Tk()
    root.withdraw()
    archivos = filedialog.askopenfilenames(
        title="Selecciona los planos .SLDDRW",
        filetypes=[("SolidWorks Drawings", "*.slddrw")]
    )
    return list(archivos)

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
    archivos = seleccionar_archivos()
    if not archivos:
        print("No se seleccionaron archivos.")
        return

    swApp = win32com.client.Dispatch('SldWorks.Application')
    swApp.Visible = False

    with tempfile.TemporaryDirectory() as temp_dir:
        pdfs = []
        for slddrw_path in archivos:
            pdf = exportar_a_pdf(swApp, slddrw_path, temp_dir)
            if pdf:
                pdfs.append(pdf)

        swApp.ExitApp()

        if not pdfs:
            print("No se generaron PDFs.")
            return

        # Guardar automáticamente el PDF combinado en la carpeta del primer archivo seleccionado
        output_folder = os.path.dirname(archivos[0])
        output_pdf = os.path.join(output_folder, "Planos_Combinados.pdf")

        # Unir los PDFs
        merger = PdfMerger()
        for pdf in pdfs:
            merger.append(pdf)
        merger.write(output_pdf)
        merger.close()
        print(f"PDF combinado guardado en: {output_pdf}")
        try:
            import tkinter.messagebox as messagebox
            messagebox.showinfo("Éxito", f"PDF combinado guardado en:\n{output_pdf}")
        except Exception:
            pass

if __name__ == "__main__":
    main()
