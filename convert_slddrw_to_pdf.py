

import os
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger
import tempfile
import threading
from stl import mesh
import numpy as np
import trimesh

# Configurar matplotlib para usar aceleración por hardware
import matplotlib
matplotlib.use('TkAgg')  # Backend optimizado
import matplotlib.pyplot as plt
from mpl_toolkits import mplot3d
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Habilitar aceleración de hardware en matplotlib
plt.rcParams['path.simplify'] = True
plt.rcParams['path.simplify_threshold'] = 1.0
plt.rcParams['agg.path.chunksize'] = 10000

# Desactivar animaciones para cambios instantáneos
plt.rcParams['animation.html'] = 'none'
matplotlib.rcParams['toolbar'] = 'None'

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertir SLDDRW a PDF y Visualizar SLDPRT")
        self.root.geometry("900x700")
        self.archivos = []
        self.temp_files = []  # Lista para rastrear archivos temporales
        
        # Configurar evento de cierre para limpiar archivos temporales
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Frame principal con dos secciones
        self.frame_pdf = tk.LabelFrame(root, text="Convertir SLDDRW a PDF", padx=10, pady=10)
        self.frame_pdf.pack(padx=10, pady=10, fill=tk.X)

        self.btn_seleccionar = tk.Button(self.frame_pdf, text="Seleccionar archivos .SLDDRW", command=self.seleccionar_archivos)
        self.btn_seleccionar.grid(row=0, column=0, pady=5, sticky="ew")

        self.btn_convertir = tk.Button(self.frame_pdf, text="Convertir a PDF", command=self.iniciar_conversion, state="disabled")
        self.btn_convertir.grid(row=1, column=0, pady=5, sticky="ew")

        self.progress = ttk.Progressbar(self.frame_pdf, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=2, column=0, pady=10)

        self.status = tk.Label(self.frame_pdf, text="Seleccione archivos para comenzar.")
        self.status.grid(row=3, column=0, pady=5)

        # Frame para visualización 3D
        self.frame_3d = tk.LabelFrame(root, text="Visualizar Pieza 3D (.SLDPRT)", padx=10, pady=10)
        self.frame_3d.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Frame superior con controles
        self.frame_controles = tk.Frame(self.frame_3d)
        self.frame_controles.pack(fill=tk.X, pady=5)
        
        # Botón para cargar archivo SLDPRT
        self.btn_abrir_3d = tk.Button(self.frame_controles, text="Abrir .SLDPRT", command=self.abrir_sldprt, bg='#4CAF50', fg='white')
        self.btn_abrir_3d.pack(side=tk.LEFT, padx=5)
        
        # Botones de vista preestablecida
        tk.Button(self.frame_controles, text="Frontal", command=lambda: self.cambiar_vista(0, 0)).pack(side=tk.LEFT, padx=2)
        tk.Button(self.frame_controles, text="Superior", command=lambda: self.cambiar_vista(90, 0)).pack(side=tk.LEFT, padx=2)
        tk.Button(self.frame_controles, text="Lateral", command=lambda: self.cambiar_vista(0, 90)).pack(side=tk.LEFT, padx=2)
        tk.Button(self.frame_controles, text="Isométrica", command=lambda: self.cambiar_vista(30, 45)).pack(side=tk.LEFT, padx=2)
        tk.Button(self.frame_controles, text="Reset Zoom", command=self.restablecer_zoom).pack(side=tk.LEFT, padx=2)
        
        # Botón para exportar plano técnico
        tk.Button(self.frame_controles, text="Exportar Plano 2D", command=self.exportar_plano_tecnico, bg='#2196F3', fg='white').pack(side=tk.LEFT, padx=5)
        
        # Checkbox para modo wireframe
        self.wireframe_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.frame_controles, text="Wireframe", variable=self.wireframe_var, command=self.actualizar_visualizacion).pack(side=tk.LEFT, padx=5)
        
        # Botón de calibración
        tk.Button(self.frame_controles, text="⚙ Calibrar", command=self.mostrar_calibracion, bg='#FF9800', fg='white').pack(side=tk.LEFT, padx=5)

        self.label_3d = tk.Label(self.frame_3d, text="Ningún archivo 3D cargado")
        self.label_3d.pack(pady=5)

        self.frame_viewer = tk.Frame(self.frame_3d)
        self.frame_viewer.pack(fill=tk.BOTH, expand=True)

        # Variables para almacenar el estado de visualización
        self.current_stl_path = None
        self.current_ax = None
        self.current_figura = None
        self.current_trimesh = None  # Almacenar mesh procesado
        self.canvas_3d = None  # Canvas de matplotlib para 3D
        self.aristas_tecnicas_cache = None  # Cache de aristas extraídas
        
        # Estado persistente de la vista (compartido entre wireframe y sólido)
        self.vista_actual = {'elev': 30, 'azim': 45, 'xlim': None, 'ylim': None, 'zlim': None}
        
        # Calibración de compensación entre modos (ajustable por usuario)
        self.calibracion = {'elev_offset': 0.0, 'azim_offset': 0.0, 'zoom_offset': 1.0}
        
        # Ventana de calibración (se crea cuando se necesita)
        self.ventana_calibracion = None
    
    def mostrar_calibracion(self):
        """Mostrar ventana de calibración de vistas"""
        if self.ventana_calibracion and self.ventana_calibracion.winfo_exists():
            self.ventana_calibracion.lift()
            return
        
        self.ventana_calibracion = tk.Toplevel(self.root)
        self.ventana_calibracion.title("Calibración de Vista Wireframe")
        self.ventana_calibracion.geometry("400x250")
        
        tk.Label(self.ventana_calibracion, text="Ajustar compensación entre modos", 
                font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Frame para controles
        frame_ajustes = tk.Frame(self.ventana_calibracion)
        frame_ajustes.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Elevación offset
        tk.Label(frame_ajustes, text="Compensación Elevación (°):").grid(row=0, column=0, sticky='w', pady=5)
        self.elev_offset_var = tk.DoubleVar(value=self.calibracion['elev_offset'])
        elev_scale = tk.Scale(frame_ajustes, from_=-10, to=10, resolution=0.1, orient=tk.HORIZONTAL,
                             variable=self.elev_offset_var, command=self.aplicar_calibracion, length=200)
        elev_scale.grid(row=0, column=1, pady=5)
        
        # Azimut offset
        tk.Label(frame_ajustes, text="Compensación Azimut (°):").grid(row=1, column=0, sticky='w', pady=5)
        self.azim_offset_var = tk.DoubleVar(value=self.calibracion['azim_offset'])
        azim_scale = tk.Scale(frame_ajustes, from_=-10, to=10, resolution=0.1, orient=tk.HORIZONTAL,
                             variable=self.azim_offset_var, command=self.aplicar_calibracion, length=200)
        azim_scale.grid(row=1, column=1, pady=5)
        
        # Zoom offset
        tk.Label(frame_ajustes, text="Factor Zoom:").grid(row=2, column=0, sticky='w', pady=5)
        self.zoom_offset_var = tk.DoubleVar(value=self.calibracion['zoom_offset'])
        zoom_scale = tk.Scale(frame_ajustes, from_=0.8, to=1.2, resolution=0.01, orient=tk.HORIZONTAL,
                             variable=self.zoom_offset_var, command=self.aplicar_calibracion, length=200)
        zoom_scale.grid(row=2, column=1, pady=5)
        
        # Botones
        frame_botones = tk.Frame(self.ventana_calibracion)
        frame_botones.pack(pady=10)
        
        tk.Button(frame_botones, text="Resetear", command=self.resetear_calibracion).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_botones, text="Aplicar y Cerrar", command=self.ventana_calibracion.destroy).pack(side=tk.LEFT, padx=5)
    
    def aplicar_calibracion(self, event=None):
        """Aplicar valores de calibración en tiempo real"""
        self.calibracion['elev_offset'] = self.elev_offset_var.get()
        self.calibracion['azim_offset'] = self.azim_offset_var.get()
        self.calibracion['zoom_offset'] = self.zoom_offset_var.get()
        
        # Si hay una vista activa, actualizar inmediatamente
        if self.current_stl_path and self.wireframe_var.get():
            self.actualizar_visualizacion()
    
    def resetear_calibracion(self):
        """Resetear calibración a valores por defecto"""
        self.elev_offset_var.set(0.0)
        self.azim_offset_var.set(0.0)
        self.zoom_offset_var.set(1.0)
        self.aplicar_calibracion()
    
    def extraer_caracteristicas_tecnicas(self, stl_path):
        """
        Extrae características geométricas usando trimesh para wireframe técnico
        Detecta: aristas agudas, bordes, características cilíndricas, planas
        """
        # Cargar mesh con trimesh para análisis avanzado
        tm = trimesh.load(stl_path)
        
        # 1. Detectar aristas por ángulo de características (sharp features)
        # Ángulo de 30 grados para detectar aristas técnicas importantes
        aristas_agudas = tm.edges[trimesh.grouping.group_rows(
            tm.edges_sorted, require_count=1)]
        
        # 2. Identificar aristas con ángulo diedro significativo
        # Esto captura cambios de dirección importantes para planos técnicos
        angulos_diedros = tm.face_adjacency_angles
        umbral_angulo = np.radians(20)  # 20 grados - aristas técnicas
        
        # Obtener aristas con ángulo significativo
        aristas_tecnicas = []
        if len(tm.face_adjacency) > 0:
            for i, angulo in enumerate(angulos_diedros):
                if abs(angulo) > umbral_angulo:
                    # Obtener la arista compartida entre las caras adyacentes
                    face_pair = tm.face_adjacency[i]
                    cara1 = tm.faces[face_pair[0]]
                    cara2 = tm.faces[face_pair[1]]
                    
                    # Encontrar vértices compartidos
                    vertices_compartidos = np.intersect1d(cara1, cara2)
                    if len(vertices_compartidos) == 2:
                        v1, v2 = tm.vertices[vertices_compartidos]
                        aristas_tecnicas.append((tuple(v1), tuple(v2)))
        
        # 3. Detectar bordes/silueta del modelo
        aristas_borde = []
        edges_unique = trimesh.grouping.group_rows(tm.edges_sorted, require_count=1)
        for edge_idx in edges_unique:
            edge = tm.edges[edge_idx]
            v1, v2 = tm.vertices[edge]
            aristas_borde.append((tuple(v1), tuple(v2)))
        
        # Combinar todas las características detectadas
        todas_aristas = set(aristas_tecnicas + aristas_borde)
        
        return list(todas_aristas), tm

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

    def abrir_sldprt(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona un archivo .SLDPRT",
            filetypes=[("SolidWorks Parts", "*.sldprt")]
        )
        if archivo:
            self.label_3d.config(text=f"Cargando: {os.path.basename(archivo)}")
            self.root.update()
            threading.Thread(target=self.cargar_y_visualizar, args=(archivo,)).start()

    def cargar_y_visualizar(self, sldprt_path):
        try:
            swApp = win32com.client.Dispatch('SldWorks.Application')
            swApp.Visible = False

            # Usar carpeta temporal del sistema pero mantener el archivo
            temp_dir = tempfile.gettempdir()
            stl_path = self.convertir_a_stl(swApp, sldprt_path, temp_dir)
            
            swApp.ExitApp()
            
            if stl_path:
                # Ejecutar visualización en el hilo principal
                self.root.after(0, self.visualizar_stl, stl_path)
                self.label_3d.config(text=f"Visualizando: {os.path.basename(sldprt_path)}")
            else:
                self.label_3d.config(text="Error al convertir el archivo")
                self.root.after(0, messagebox.showerror, "Error", "No se pudo convertir el archivo a STL")

        except Exception as e:
            self.label_3d.config(text="Error al cargar el archivo")
            self.root.after(0, messagebox.showerror, "Error", f"Error al cargar el archivo:\n{e}")

    def convertir_a_stl(self, swApp, sldprt_path, temp_dir):
        from win32com.client import VARIANT
        import pythoncom

        swDocPART = 1
        swOpenDocOptions_Silent = 64

        errors = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        part = swApp.OpenDoc6(sldprt_path, swDocPART, swOpenDocOptions_Silent, '', errors, warnings)
        if part is None:
            print(f"No se pudo abrir {os.path.basename(sldprt_path)}")
            return None

        stl_path = os.path.join(temp_dir, os.path.splitext(os.path.basename(sldprt_path))[0] + '.stl')

        try:
            result = part.SaveAs(stl_path)
            if result:
                print(f"STL guardado: {stl_path}")
                self.temp_files.append(stl_path)  # Rastrear archivo temporal
                return stl_path
            else:
                print(f"Error al guardar STL: {stl_path}")
                return None
        except Exception as e:
            print(f"Error exportando a STL: {e}")
            return None
        finally:
            swApp.CloseDoc(os.path.basename(sldprt_path))

    def visualizar_stl(self, stl_path, restaurar_vista=None):
        # Cerrar canvas anterior si existe
        if self.canvas_3d:
            self.canvas_3d.get_tk_widget().destroy()
        
        # Cerrar figura anterior para liberar memoria
        if self.current_figura:
            plt.close(self.current_figura)

        self.current_stl_path = stl_path
        stl_mesh = mesh.Mesh.from_file(stl_path)

        # Simplificar modelo para mejor rendimiento de rotación
        max_triangles = 2000
        if len(stl_mesh.vectors) > max_triangles:
            step = max(1, len(stl_mesh.vectors) // max_triangles)
            simplified_vectors = stl_mesh.vectors[::step]
        else:
            simplified_vectors = stl_mesh.vectors

        # PRE-PROCESAR: Extraer aristas técnicas ANTES de crear figura (evita interferencia)
        aristas_tecnicas = None
        if self.wireframe_var.get():
            if self.current_trimesh is None or self.current_stl_path != stl_path:
                aristas_tecnicas, self.current_trimesh = self.extraer_caracteristicas_tecnicas(stl_path)
                self.aristas_tecnicas_cache = aristas_tecnicas  # Cachear aristas
            else:
                aristas_tecnicas = self.aristas_tecnicas_cache

        # Crear figura optimizada DESPUÉS del pre-procesamiento
        self.current_figura = plt.figure(figsize=(8, 6), dpi=80)
        self.current_ax = self.current_figura.add_subplot(111, projection='3d')
        
        # Modo wireframe o sólido
        if self.wireframe_var.get():
            # Modo wireframe AVANZADO - usar aristas pre-calculadas

            for arista in aristas_tecnicas:
                p1, p2 = arista
                self.current_ax.plot3D([p1[0], p2[0]], [p1[1], p2[1]], [p1[2], p2[2]], 
                                       color='#1a1a1a', linewidth=1.5, alpha=0.9)
            
            collection = None  # No hay colección en modo wireframe
        else:
            # Modo sólido con sombreado mejorado
            # Calcular normales para sombreado realista
            normales = np.array([np.cross(v[1] - v[0], v[2] - v[0]) for v in simplified_vectors])
            mag = np.linalg.norm(normales, axis=1)
            mag[mag == 0] = 1  # Evitar división por cero
            normales = normales / mag[:, np.newaxis]
            
            # Múltiples fuentes de luz para sombreado más realista
            luz_principal = np.array([1, 1, 2])  # Luz principal desde arriba-derecha
            luz_principal = luz_principal / np.linalg.norm(luz_principal)
            
            luz_relleno = np.array([-0.5, -0.5, 1])  # Luz de relleno suave
            luz_relleno = luz_relleno / np.linalg.norm(luz_relleno)
            
            # Calcular intensidad combinada
            intensidad_principal = np.abs(np.dot(normales, luz_principal))
            intensidad_relleno = np.abs(np.dot(normales, luz_relleno))
            
            # Combinar luces: 70% principal + 30% relleno + luz ambiente
            intensidades = 0.2 + 0.7 * intensidad_principal + 0.3 * intensidad_relleno
            intensidades = np.clip(intensidades, 0.3, 1.0)
            
            # Crear colores basados en intensidad con gradiente más rico
            color_base = np.array([0.6, 0.7, 0.85])  # Azul grisáceo
            color_sombra = np.array([0.3, 0.35, 0.45])  # Azul oscuro para sombras
            
            # Interpolación de color basada en intensidad
            colores = color_sombra + (color_base - color_sombra) * intensidades[:, np.newaxis]
            
            # Renderizar con aristas sutiles del mismo color (invisibles)
            collection = mplot3d.art3d.Poly3DCollection(
                simplified_vectors,
                facecolors=colores,
                edgecolors=colores,
                linewidths=0.1,
                alpha=1.0,
                antialiased=True
            )
        
        # Solo agregar colección si existe (modo sólido)
        if collection is not None:
            self.current_ax.add_collection3d(collection)
        
        # Configurar límites del gráfico (solo si no hay vista para restaurar)
        if not restaurar_vista or restaurar_vista.get('xlim') is None:
            scale = stl_mesh.points.flatten()
            self.current_ax.auto_scale_xyz(scale, scale, scale)
        
        # Ocultar ejes para vista limpia
        self.current_ax.set_axis_off()
        
        # Fondo blanco
        self.current_ax.set_facecolor('white')
        self.current_figura.patch.set_facecolor('white')
        
        # Establecer vista ANTES de crear el canvas
        if restaurar_vista:
            # Aplicar calibración si está en modo wireframe
            if self.wireframe_var.get():
                elev_ajustado = restaurar_vista['elev'] + self.calibracion['elev_offset']
                azim_ajustado = restaurar_vista['azim'] + self.calibracion['azim_offset']
                zoom_factor = self.calibracion['zoom_offset']
                
                self.current_ax.view_init(elev=elev_ajustado, azim=azim_ajustado)
                
                if restaurar_vista.get('xlim') is not None:
                    # Aplicar factor de zoom
                    xlim = restaurar_vista['xlim']
                    ylim = restaurar_vista['ylim']
                    zlim = restaurar_vista['zlim']
                    
                    # Calcular centro
                    x_center = (xlim[0] + xlim[1]) / 2
                    y_center = (ylim[0] + ylim[1]) / 2
                    z_center = (zlim[0] + zlim[1]) / 2
                    
                    # Aplicar zoom
                    x_range = (xlim[1] - xlim[0]) * zoom_factor
                    y_range = (ylim[1] - ylim[0]) * zoom_factor
                    z_range = (zlim[1] - zlim[0]) * zoom_factor
                    
                    self.current_ax.set_xlim([x_center - x_range/2, x_center + x_range/2])
                    self.current_ax.set_ylim([y_center - y_range/2, y_center + y_range/2])
                    self.current_ax.set_zlim([z_center - z_range/2, z_center + z_range/2])
            else:
                # Modo sólido - sin calibración
                self.current_ax.view_init(elev=restaurar_vista['elev'], azim=restaurar_vista['azim'])
                if restaurar_vista.get('xlim') is not None:
                    self.current_ax.set_xlim(restaurar_vista['xlim'])
                    self.current_ax.set_ylim(restaurar_vista['ylim'])
                    self.current_ax.set_zlim(restaurar_vista['zlim'])
        else:
            # Vista isométrica por defecto para carga inicial
            self.current_ax.view_init(elev=30, azim=45)
        
        # Integrar en Tkinter
        self.canvas_3d = FigureCanvasTkAgg(self.current_figura, master=self.frame_viewer)
        self.canvas_3d.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Forzar actualización inmediata sin animación
        self.canvas_3d.draw_idle()
        self.root.update_idletasks()
        
        # Conectar eventos del mouse
        self.canvas_3d.mpl_connect('scroll_event', self.on_scroll)
        self.canvas_3d.mpl_connect('button_release_event', self.on_mouse_release)
        
        # Actualizar estado persistente con la vista actual
        self.sincronizar_vista_actual()
    
    def on_mouse_release(self, event):
        """Capturar cuando el usuario suelta el mouse después de rotar"""
        # Sincronizar vista después de cualquier interacción con el mouse
        if event.button in [1, 3]:  # Botón izquierdo o derecho
            self.root.after(50, self.sincronizar_vista_actual)  # Pequeño delay para que matplotlib actualice
    
    def sincronizar_vista_actual(self):
        """Guardar el estado actual de la vista en vista_actual"""
        if self.current_ax:
            self.vista_actual['elev'] = self.current_ax.elev
            self.vista_actual['azim'] = self.current_ax.azim
            self.vista_actual['xlim'] = self.current_ax.get_xlim()
            self.vista_actual['ylim'] = self.current_ax.get_ylim()
            self.vista_actual['zlim'] = self.current_ax.get_zlim()
    
    def on_scroll(self, event):
        """Manejar el evento de scroll del mouse para hacer zoom"""
        if self.current_ax is None:
            return
        
        # Obtener los límites actuales
        xlim = self.current_ax.get_xlim()
        ylim = self.current_ax.get_ylim()
        zlim = self.current_ax.get_zlim()
        
        # Calcular el centro de la vista
        xdata = (xlim[0] + xlim[1]) / 2
        ydata = (ylim[0] + ylim[1]) / 2
        zdata = (zlim[0] + zlim[1]) / 2
        
        # Factor de zoom (scroll up = zoom in, scroll down = zoom out)
        if event.button == 'up':
            scale_factor = 0.9  # Zoom in (acercar)
        elif event.button == 'down':
            scale_factor = 1.1  # Zoom out (alejar)
        else:
            return
        
        # Calcular nuevos límites
        new_xlim = [xdata - (xdata - xlim[0]) * scale_factor,
                    xdata + (xlim[1] - xdata) * scale_factor]
        new_ylim = [ydata - (ydata - ylim[0]) * scale_factor,
                    ydata + (ylim[1] - ydata) * scale_factor]
        new_zlim = [zdata - (zdata - zlim[0]) * scale_factor,
                    zdata + (zlim[1] - zdata) * scale_factor]
        
        # Aplicar nuevos límites
        self.current_ax.set_xlim(new_xlim)
        self.current_ax.set_ylim(new_ylim)
        self.current_ax.set_zlim(new_zlim)
        
        # Actualizar estado persistente
        self.sincronizar_vista_actual()
        
        # Redibujar
        self.canvas_3d.draw()
    
    def cambiar_vista(self, elevacion, azimut):
        """Cambiar la vista 3D a una posición preestablecida"""
        if self.current_ax:
            self.current_ax.view_init(elev=elevacion, azim=azimut)
            
            # Actualizar estado persistente
            self.sincronizar_vista_actual()
            
            self.canvas_3d.draw()
    
    def restablecer_zoom(self):
        """Restablecer el zoom a la vista completa del modelo"""
        if self.current_ax and self.current_stl_path:
            stl_mesh = mesh.Mesh.from_file(self.current_stl_path)
            scale = stl_mesh.points.flatten()
            self.current_ax.auto_scale_xyz(scale, scale, scale)
            
            # Actualizar estado persistente
            self.sincronizar_vista_actual()
            
            self.canvas_3d.draw()
    
    def actualizar_visualizacion(self):
        """Actualizar la visualización cuando cambian los controles"""
        if self.current_stl_path:
            # Sincronizar estado antes de recrear (por si hay cambios de rotación manual)
            if self.current_ax:
                self.sincronizar_vista_actual()
            
            # Recrear la visualización usando el estado persistente
            self.visualizar_stl(self.current_stl_path, restaurar_vista=self.vista_actual)
    
    def limpiar_archivos_temporales(self):
        """Eliminar todos los archivos temporales STL creados"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    print(f"Archivo temporal eliminado: {temp_file}")
            except Exception as e:
                print(f"Error al eliminar {temp_file}: {e}")
        self.temp_files.clear()
    
    def exportar_plano_tecnico(self):
        """Exporta vistas técnicas 2D del modelo para planos normados"""
        if not self.current_stl_path or not self.current_trimesh:
            messagebox.showwarning("Advertencia", "Primero carga un modelo 3D")
            return
        
        try:
            # Solicitar carpeta de destino
            carpeta_destino = filedialog.askdirectory(title="Seleccionar carpeta para guardar planos")
            if not carpeta_destino:
                return
            
            nombre_base = os.path.splitext(os.path.basename(self.current_stl_path))[0]
            
            # Generar vistas ortográficas estándar usando trimesh
            vistas = {
                'frontal': ([0, 0, 1], [0, 1, 0]),     # Vista Z, arriba Y
                'superior': ([0, 1, 0], [1, 0, 0]),    # Vista Y, arriba X
                'lateral': ([1, 0, 0], [0, 0, 1]),     # Vista X, arriba Z
                'isometrica': ([1, 1, 1], [0, 0, 1])   # Vista isométrica
            }
            
            archivos_generados = []
            
            for nombre_vista, (direccion, arriba) in vistas.items():
                # Crear figura para esta vista
                fig = plt.figure(figsize=(10, 10), dpi=150)
                ax = fig.add_subplot(111)
                
                # Proyección 2D de las aristas técnicas
                aristas_tecnicas, _ = self.extraer_caracteristicas_tecnicas(self.current_stl_path)
                
                # Proyectar aristas a 2D según la dirección de vista
                direccion = np.array(direccion, dtype=float)
                direccion = direccion / np.linalg.norm(direccion)
                arriba = np.array(arriba, dtype=float)
                arriba = arriba / np.linalg.norm(arriba)
                
                # Crear sistema de coordenadas 2D
                derecha = np.cross(arriba, direccion)
                derecha = derecha / np.linalg.norm(derecha)
                
                # Proyectar cada arista
                for arista in aristas_tecnicas:
                    p1, p2 = np.array(arista[0]), np.array(arista[1])
                    
                    # Proyectar puntos 3D a 2D
                    p1_2d = np.array([np.dot(p1, derecha), np.dot(p1, arriba)])
                    p2_2d = np.array([np.dot(p2, derecha), np.dot(p2, arriba)])
                    
                    # Dibujar línea
                    ax.plot([p1_2d[0], p2_2d[0]], [p1_2d[1], p2_2d[1]], 
                           'k-', linewidth=0.5)
                
                # Configurar aspecto técnico
                ax.set_aspect('equal')
                ax.grid(True, alpha=0.3, linestyle='--')
                ax.set_title(f'Vista {nombre_vista.capitalize()}', fontsize=12, fontweight='bold')
                
                # Guardar como PNG de alta calidad
                archivo_salida = os.path.join(carpeta_destino, f"{nombre_base}_{nombre_vista}.png")
                plt.savefig(archivo_salida, dpi=300, bbox_inches='tight', 
                           facecolor='white', edgecolor='none')
                plt.close(fig)
                
                archivos_generados.append(archivo_salida)
                print(f"Vista {nombre_vista} guardada: {archivo_salida}")
            
            messagebox.showinfo("Éxito", 
                              f"Se generaron {len(archivos_generados)} vistas técnicas:\n" + 
                              "\n".join([os.path.basename(f) for f in archivos_generados]))
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar planos: {str(e)}")
            print(f"Error detallado: {e}")
    
    def on_closing(self):
        """Manejar el evento de cierre de la ventana"""
        print("\nCerrando aplicación y limpiando archivos temporales...")
        self.limpiar_archivos_temporales()
        self.root.destroy()


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
