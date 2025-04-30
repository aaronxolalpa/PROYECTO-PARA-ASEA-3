import os
import shutil
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import fitz  
from PIL import Image
from pyzbar.pyzbar import decode
import io
import pyodbc
import subprocess
from datetime import datetime

# Configuración de la base de datos Access
DB_PATH = "C:/Users/sergio/Desktop/guias.accdb"
TABLE_NAME = "Copia de Registro"
COLUMN_GUIA = "GUIA"
COLUMN_OFICIO = "OFICIO"
TABLE_NOTIF = "Copia de notificaciones_e2"
COLUMN_FECHA_NOTIF = "f_notif_oficio"
COLUMN_VIA_NOTIF = "via_notif"  # Nueva columna para vía de notificación

# Carpeta temporal para almacenar los archivos
TEMP_FOLDER = os.path.join(tempfile.gettempdir(), "pdf_temporales")
os.makedirs(TEMP_FOLDER, exist_ok=True)

def conectar_bd():
    """Conecta a la base de datos Access."""
    try:
        conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + DB_PATH + ";"
        conn = pyodbc.connect(conn_str)
        return conn
    except Exception as e:
        messagebox.showerror("Error de Base de Datos", f"Error al conectar con la base de datos: {e}")
        return None

def buscar_oficio_por_guia(guia):
    """Busca el número de oficio asociado a una guía en la base de datos."""
    conn = conectar_bd()
    if not conn:
        return None

    cursor = conn.cursor()
    query = f"SELECT [{COLUMN_OFICIO}] FROM [{TABLE_NAME}] WHERE [{COLUMN_GUIA}] = ?"
    
    try:
        cursor.execute(query, (guia,))
        resultado = cursor.fetchone()
        conn.close()  # Cerramos la conexión aquí
        return resultado[0] if resultado else None
    except Exception as e:
        if conn:
            conn.close()  # Aseguramos cerrar la conexión en caso de error
        messagebox.showerror("Error SQL", f"Error al ejecutar la consulta SQL: {e}")
        return None

def actualizar_fecha_notificacion(oficio, fecha_obj, tipo="CC"):
    """Actualiza la fecha de notificación y vía de notificación para un oficio específico."""
    if not oficio:
        messagebox.showerror("Error", "No hay número de oficio para actualizar.")
        return False
        
    if not fecha_obj:
        messagebox.showerror("Error", "Debe proporcionar una fecha válida.")
        return False
        
    conn = conectar_bd()
    if not conn:
        return False

    cursor = conn.cursor()
    
    # Determinar el valor para via_notif según el tipo
    via_notif = "Correo certificado" if tipo == "CC" else ""
    
    # Query para actualizar tanto la fecha como la vía de notificación
    query = f"UPDATE [{TABLE_NOTIF}] SET [{COLUMN_FECHA_NOTIF}] = ?, [{COLUMN_VIA_NOTIF}] = ? WHERE [{COLUMN_OFICIO}] = ?"
    
    try:
        # Pasamos el objeto datetime y la vía de notificación a Access
        cursor.execute(query, (fecha_obj, via_notif, oficio))
        conn.commit()
        filas_afectadas = cursor.rowcount
        conn.close()  # Cerramos la conexión aquí
        
        if filas_afectadas > 0:
            fecha_formateada = fecha_obj.strftime("%d/%m/%Y")
            mensaje = f"Se actualizó la fecha de notificación ({fecha_formateada}) para el oficio {oficio}."
            if tipo == "CC":
                mensaje += " Vía de notificación: Correo certificado."
            messagebox.showinfo("Actualización Exitosa", mensaje)
            return True
        else:
            messagebox.showwarning("Advertencia", f"No se encontró el oficio {oficio} en la tabla {TABLE_NOTIF}.")
            return False
    except Exception as e:
        if conn:
            conn.close()  # Aseguramos cerrar la conexión en caso de error
        messagebox.showerror("Error SQL", f"Error al actualizar la fecha de notificación: {e}")
        return False

def extraer_imagenes_pdf(pdf_path):
    """Extrae imágenes de un PDF."""
    imagenes = []
    try:
        pdf = fitz.open(pdf_path)
        for pagina_num in range(len(pdf)):
            pagina = pdf[pagina_num]
            imagen_list = pagina.get_images(full=True)
            for img_index, img in enumerate(imagen_list):
                xref = img[0]
                base_image = pdf.extract_image(xref)
                imagen_bytes = base_image["image"]
                imagen = Image.open(io.BytesIO(imagen_bytes))
                imagenes.append(imagen)
        return imagenes
    except Exception as e:
        messagebox.showerror("Error", f"Error al extraer imágenes del PDF: {e}")
        return []

def leer_codigos_barra(imagen):
    """Lee códigos de barras en una imagen."""
    codigos = decode(imagen)
    resultados = []
    for codigo in codigos:
        data = codigo.data.decode('utf-8')
        tipo = codigo.type
        resultados.append((tipo, data))
    return resultados

def abrir_carpeta(ruta):
    """Abre la carpeta en el explorador de archivos."""
    if os.path.exists(ruta):
        if os.name == 'nt':  # Windows
            os.startfile(ruta)
        elif os.name == 'posix':  # macOS, Linux
            subprocess.call(['open', ruta] if os.name == 'darwin' else ['xdg-open', ruta])
    else:
        messagebox.showerror("Error", f"La carpeta {ruta} no existe.")

def abrir_pdf(pdf_path):
    """Abre un archivo PDF con el visor predeterminado del sistema."""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(pdf_path)
        elif os.name == 'posix':  # macOS, Linux
            subprocess.call(['open', pdf_path] if os.name == 'darwin' else ['xdg-open', pdf_path])
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Error al abrir el PDF: {e}")
        return False

def cerrar_pdf():
    """Intenta cerrar el visor de PDF actualmente abierto."""
    try:
        if os.name == 'nt':  # Windows
            # Intenta cerrar Adobe Reader, Acrobat, Edge, etc.
            aplicaciones = ['AcroRd32.exe', 'Acrobat.exe', 'msedge.exe', 'chrome.exe', 'firefox.exe']
            for app in aplicaciones:
                try:
                    subprocess.run(['taskkill', '/f', '/im', app], 
                                  stdout=subprocess.DEVNULL, 
                                  stderr=subprocess.DEVNULL)
                except:
                    pass
        # Para macOS y Linux, no cerramos aplicaciones ya que podría afectar a otras ventanas
        # que el usuario tenga abiertas
        return True
    except Exception as e:
        print(f"Error al intentar cerrar el PDF: {e}")
        return False

# función para validar fechas
def validar_fecha(dia, mes, anio):
    
    # Validar que todos los campos sean números
    try:
        dia_num = int(dia)
        mes_num = int(mes)
        anio_num = int(anio)
    except ValueError:
        return False, "Los campos de fecha deben contener solo números."
    
    # Validar rangos básicos
    if mes_num < 1 or mes_num > 12:
        return False, f"El mes debe estar entre 1 y 12. Valor actual: {mes_num}"
    
    if dia_num < 1:
        return False, f"El día debe ser mayor a 0. Valor actual: {dia_num}"
    
    # Obtener el último día del mes
    if mes_num in [4, 6, 9, 11]:  # Abril, Junio, Septiembre, Noviembre
        ultimo_dia = 30
    elif mes_num == 2:  # Febrero
        # Verificar si es año bisiesto
        if (anio_num % 4 == 0 and anio_num % 100 != 0) or (anio_num % 400 == 0):
            ultimo_dia = 29
        else:
            ultimo_dia = 28
    else:
        ultimo_dia = 31
    
    if dia_num > ultimo_dia:
        return False, f"El día {dia_num} no es válido para el mes {mes_num}. Último día: {ultimo_dia}"
    
    # Validar que el año tenga sentido (opcional, ajusta según necesidades)
    if anio_num < 2000 or anio_num > 2100:
        return False, f"El año {anio_num} está fuera del rango permitido (2000-2100)"
    
    # Si pasó todas las validaciones
    return True, None

class RenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Renombrar Oficios Escaneados")
        self.root.geometry("600x500")
        
        # Lista de archivos PDF a procesar
        self.pdf_files = []
        self.skipped_files = []
        self.current_pdf_index = -1
        self.pdf_viewer_process = None
        
        # Contador de archivos
        self.contador_frame = tk.Frame(root)
        self.contador_frame.pack(pady=5, fill="x")
        
        self.files_loaded_label = tk.Label(self.contador_frame, text="Archivos cargados: 0")
        self.files_loaded_label.pack(side="left", padx=10)
        
        self.files_processed_label = tk.Label(self.contador_frame, text="Archivos procesados: 0/0")
        self.files_processed_label.pack(side="right", padx=10)
        
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=580, mode="determinate")
        self.progress_bar.pack(pady=5, padx=10)
        
        self.rename_frame = tk.LabelFrame(root, text="Generar nombre a partir N° de oficio y fecha")
        self.rename_frame.pack(pady=10, fill="both", expand=True)
        
        button_frame = tk.Frame(self.rename_frame)
        button_frame.pack(pady=5)
        
        
        self.folder_button = tk.Button(button_frame, text="Seleccionar Carpeta", command=self.seleccionar_carpeta)
        self.folder_button.pack(side="left", padx=5)
        
        self.save_button = tk.Button(button_frame, text="Guardar PDF", command=self.guardar_pdf)
        self.save_button.pack(side="left", padx=5)
    
        self.open_folder_button = tk.Button(button_frame, text="Abrir Carpeta de PDF´s renombrados", command=lambda: abrir_carpeta(TEMP_FOLDER))
        self.open_folder_button.pack(side="left", padx=5)
        
        tipo_frame = tk.Frame(self.rename_frame)
        tipo_frame.pack(pady=5)
        
        self.tipo_label = tk.Label(tipo_frame, text="Tipo:")
        self.tipo_label.pack(side="left", padx=5)
        
        self.tipo_var = tk.StringVar(value="CC")
        self.cc_radio = tk.Radiobutton(tipo_frame, text="CC", variable=self.tipo_var, value="CC", command=self.actualizar_tipo)
        self.cc_radio.pack(side="left", padx=5)
        
        self.dv_radio = tk.Radiobutton(tipo_frame, text="DEV", variable=self.tipo_var, value="DEV", command=self.actualizar_tipo)
        self.dv_radio.pack(side="left", padx=5)
        
        date_frame = tk.Frame(self.rename_frame)
        date_frame.pack(pady=5)
        
        # Día
        self.dia_label = tk.Label(date_frame, text="Día")
        self.dia_label.pack(side="left", padx=5)
        self.dia_entry = tk.Entry(date_frame, width=10)
        self.dia_entry.pack(side="left", padx=5)

        # Mes
        self.mes_label = tk.Label(date_frame, text="Mes")
        self.mes_label.pack(side="left", padx=5)
        self.mes_entry = tk.Entry(date_frame, width=10)
        self.mes_entry.pack(side="left", padx=5)

        # Año - Añadimos valor predeterminado "20"
        self.anio_label = tk.Label(date_frame, text="Año")
        self.anio_label.pack(side="left", padx=5)
        self.anio_entry = tk.Entry(date_frame, width=10)
        self.anio_entry.insert(0, "20")  # Valor predeterminado "20"
        self.anio_entry.pack(side="left", padx=5)
        
        # Botón para actualizar fecha de notificación
        self.update_date_button = tk.Button(date_frame, text="Actualizar Fecha", command=self.actualizar_fecha_bd)
        self.update_date_button.pack(side="left", padx=5)
        
        guia_frame = tk.Frame(self.rename_frame)
        guia_frame.pack(pady=5)
        
        self.guia_label = tk.Label(guia_frame, text="Guía Escaneada")
        self.guia_label.pack(side="left", padx=5)
        
        self.guia_entry = tk.Entry(guia_frame, width=30)
        self.guia_entry.pack(side="left", padx=5)
        
        oficio_frame = tk.Frame(self.rename_frame)
        oficio_frame.pack(pady=5)
        
        self.oficio_label = tk.Label(oficio_frame, text="Número de Oficio")
        self.oficio_label.pack(side="left", padx=5)
        
        self.oficio_entry = tk.Entry(oficio_frame, width=30)
        self.oficio_entry.pack(side="left", padx=5)
        
        self.filename_label = tk.Label(self.rename_frame, text="Nombre de archivo:")
        self.filename_label.pack()
        
        self.filename_entry = tk.Entry(self.rename_frame, width=50)
        self.filename_entry.pack()
        
        # Frame para mostrar archivos saltados
        self.skipped_frame = tk.LabelFrame(root, text="Archivos No PDF (saltados)")
        self.skipped_frame.pack(pady=5, fill="x", padx=10)
        
        self.skipped_text = tk.Text(self.skipped_frame, height=3, width=70)
        self.skipped_text.pack(pady=5, padx=5)
        self.skipped_text.config(state="disabled")
         
        self.exit_button = tk.Button(root, text="Salir", command=self.salir)
        self.exit_button.pack(pady=10)
        
        # Establecer el estado inicial de los campos de fecha según el radiobutton inicial
        self.actualizar_tipo()
        
        # Contadores de archivos
        self.processed_count = 0
        
    def salir(self):
        """Cierra la aplicación y cualquier PDF abierto."""
        cerrar_pdf()
        self.root.quit()
        
    def actualizar_fecha_bd(self):
        """Actualiza la fecha de notificación en la base de datos."""
        tipo = self.tipo_var.get()
        if tipo == "DEV":
            messagebox.showinfo("Información", "La actualización de fecha solo está disponible para tipo CC.")
            return
            
        dia = self.dia_entry.get().strip()
        mes = self.mes_entry.get().strip()
        anio = self.anio_entry.get().strip()
        oficio = self.oficio_entry.get().strip()
        
        # Validar que se hayan completado los campos
        if not dia or not mes or not anio:
            messagebox.showerror("Error", "Debe completar todos los campos de fecha (día, mes, año).")
            return
            
        if not oficio:
            messagebox.showerror("Error", "No hay número de oficio para actualizar.")
            return
        
        # Validar la fecha usando la nueva función
        es_valida, mensaje_error = validar_fecha(dia, mes, anio)
        if not es_valida:
            messagebox.showerror("Error de Fecha", mensaje_error)
            return
        
        # Si la fecha es válida, continuamos con el proceso
        try:
            dia_num = int(dia)
            mes_num = int(mes)
            anio_num = int(anio)
            
            # Crear objeto datetime
            fecha_obj = datetime(anio_num, mes_num, dia_num)
            
            # Actualizar en la base de datos - pasamos el tipo seleccionado
            if actualizar_fecha_notificacion(oficio, fecha_obj, tipo):
                # Si se actualizó correctamente, actualizar también el nombre del archivo
                self.actualizar_nombre()
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar la fecha: {e}")
    
    def seleccionar_carpeta(self):
        """Selecciona una carpeta y carga todos los archivos PDF dentro de ella."""
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return
        
        # Cerrar cualquier PDF abierto previamente
        cerrar_pdf()
        
        self.pdf_files = []
        self.skipped_files = []
        
        # Limpiar texto de archivos saltados
        self.skipped_text.config(state="normal")
        self.skipped_text.delete(1.0, tk.END)
        self.skipped_text.config(state="disabled")
        
        # Buscar todos los archivos en la carpeta
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                if filename.lower().endswith('.pdf'):
                    self.pdf_files.append(file_path)
                else:
                    self.skipped_files.append(filename)
        
        # Actualizar lista de archivos saltados
        if self.skipped_files:
            self.skipped_text.config(state="normal")
            self.skipped_text.insert(tk.END, ", ".join(self.skipped_files))
            self.skipped_text.config(state="disabled")
        
        self.current_pdf_index = -1
        self.processed_count = 0
        
        total_files = len(self.pdf_files)
        self.files_loaded_label.config(text=f"Archivos cargados: {total_files}")
        self.files_processed_label.config(text=f"Archivos procesados: 0/{total_files}")
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        
        if total_files > 0:
            messagebox.showinfo("Información", f"Se encontraron {total_files} archivos PDF en la carpeta.\nSe omitieron {len(self.skipped_files)} archivos que no son PDF.")
            # Procesar el primer archivo
            self.procesar_siguiente()
        else:
            messagebox.showwarning("Advertencia", "No se encontraron archivos PDF en la carpeta seleccionada.")
        
    def procesar_siguiente(self):
        """Procesa el siguiente archivo PDF de la lista."""
        if not self.pdf_files:
            messagebox.showinfo("Información", "No hay archivos PDF cargados.")
            return
            
        self.current_pdf_index += 1
        
        if self.current_pdf_index >= len(self.pdf_files):
            messagebox.showinfo("Proceso Completado", "Todos los archivos PDF han sido procesados.")
            self.current_pdf_index = len(self.pdf_files) - 1
            return
            
        # Cerrar cualquier PDF abierto previamente
        cerrar_pdf()
        
        # Limpiar campos
        self.limpiar_campos()
        
        # Cargar el archivo actual
        current_pdf = self.pdf_files[self.current_pdf_index]
        self.procesar_pdf(current_pdf)
        
        # Abrir el PDF actual
        self.root.after(500, lambda: abrir_pdf(current_pdf))
        
        # Actualizar etiquetas
        self.files_processed_label.config(text=f"Archivos procesados: {self.processed_count}/{len(self.pdf_files)}")
        
    def procesar_pdf(self, pdf_path):
        """Procesa un archivo PDF."""
        self.pdf_path = pdf_path
        self.filename_entry.delete(0, tk.END)
        self.filename_entry.insert(0, os.path.basename(pdf_path))
        
        imagenes = extraer_imagenes_pdf(pdf_path)
        
        if not imagenes:
            messagebox.showerror("Error", f"No se encontraron imágenes en el PDF: {os.path.basename(pdf_path)}")
            return
        
        for imagen in imagenes:
            codigos = leer_codigos_barra(imagen)
            if codigos:
                for _, data in codigos:
                    self.guia_entry.delete(0, tk.END)
                    self.guia_entry.insert(0, data)
                    numero_oficio = buscar_oficio_por_guia(data)
                    if numero_oficio:
                        self.oficio_entry.delete(0, tk.END)
                        self.oficio_entry.insert(0, numero_oficio)
                        self.actualizar_nombre()
                        return
        
        messagebox.showwarning("Atención", f"No se encontró un número de oficio para las guías escaneadas en el archivo: {os.path.basename(pdf_path)}")
    
    def actualizar_tipo(self):
        """Actualiza el estado de los campos de fecha según el tipo seleccionado y actualiza el nombre."""
        tipo = self.tipo_var.get()
        if tipo == "DEV":
            # Deshabilitar campos de fecha
            self.dia_entry.config(state="disabled")
            self.mes_entry.config(state="disabled")
            self.anio_entry.config(state="disabled")
            self.update_date_button.config(state="disabled")
        else:
            # Habilitar campos de fecha
            self.dia_entry.config(state="normal")
            self.mes_entry.config(state="normal")
            self.anio_entry.config(state="normal")
            self.update_date_button.config(state="normal")
        
        # Actualizar el nombre basado en el nuevo tipo
        self.actualizar_nombre()
        
    def actualizar_nombre(self):
        numero_oficio = self.oficio_entry.get().strip().replace("/", "_")
        tipo = self.tipo_var.get().lower()
        if numero_oficio:
            nombre_archivo = f"{numero_oficio}_{tipo}.pdf"
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, nombre_archivo)
    
    def limpiar_campos(self):
        """Limpia todos los campos del formulario."""
        self.filename_entry.delete(0, tk.END)
        self.guia_entry.delete(0, tk.END)
        self.oficio_entry.delete(0, tk.END)
        self.dia_entry.delete(0, tk.END)
        self.mes_entry.delete(0, tk.END)
        # Limpiamos el año y volvemos a poner el valor predeterminado "20"
        self.anio_entry.delete(0, tk.END)
        self.anio_entry.insert(0, "20")
        
    def guardar_pdf(self):
        """Guarda el PDF actual, lo cierra y avanza al siguiente."""
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            messagebox.showerror("Error", "No hay un PDF cargado para guardar.")
            return
        
        nombre_archivo = self.filename_entry.get().strip()
        if not nombre_archivo:
            messagebox.showerror("Error", "Debe proporcionar un nombre para el archivo.")
            return
        
        # Cerrar cualquier PDF abierto
        cerrar_pdf()
        
        destino = os.path.join(TEMP_FOLDER, nombre_archivo)
        
        try:
            shutil.copy(self.pdf_path, destino)
            messagebox.showinfo("Guardado Exitoso", f"El archivo se guardó en: {destino}")
            
            # Incrementar contador de archivos procesados
            self.processed_count += 1
            self.files_processed_label.config(text=f"Archivos procesados: {self.processed_count}/{len(self.pdf_files)}")
            self.progress_bar["value"] = self.processed_count
            
            # Procesar el siguiente archivo
            self.procesar_siguiente()
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = RenameApp(root)
    root.mainloop()