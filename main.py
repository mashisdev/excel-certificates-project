import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time

# Lógica de procesamiento de Excel
def resumir_certificados(ruta_archivo, nombre_hoja='Hoja1'):
    """
    Agrupa registros de un archivo de Excel por 'certificado' y consolida los datos.
    Suma solo los valores únicos de las columnas de montos.
    
    Args:
        ruta_archivo (str): La ruta al archivo de Excel a procesar.
        nombre_hoja (str): El nombre de la hoja de cálculo.
    """
    try:
        # Simulación de un proceso largo
        time.sleep(2) 
        
        # Cargar el archivo de Excel en un DataFrame de pandas
        dtype_cols = {
            'sec_ejec': str,
            'certificado': str,
            'certificacion': str,
            'secuencia': str,
            'expediente': str
        }
        df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, dtype=dtype_cols)
        
        df.columns = df.columns.str.strip()
        
        def sum_unique(series):
            return series.drop_duplicates().sum()

        agregaciones = {
            'ano_eje': 'first',
            'sec_ejec': 'first',
            'certificacion': 'first',
            'secuencia': 'first',
            'ciclo': 'first',
            'fase': 'first',
            'moneda': lambda x: ', '.join(x.dropna().unique()),
            'expediente': 'first',
            'monto_certificado': sum_unique,
            'monto_comp_anual': sum_unique,
            'compromiso': sum_unique,
            'devengado': sum_unique,
            'girado': sum_unique,
        }
        
        df_resumido = df.groupby('certificado', as_index=False).agg(agregaciones)
        df_resumido = df_resumido.sort_values(by='certificado')
        
        nombre_archivo_salida = 'certificados_resumidos.xlsx'
        df_resumido.to_excel(nombre_archivo_salida, index=False)
        
        return f"Proceso completado. Los datos consolidados se han guardado en '{nombre_archivo_salida}'.", "success"
        
    except FileNotFoundError:
        return "Error: El archivo no se encontró.", "error"
    except KeyError as e:
        return f"Error: La columna {e} no se encontró en el archivo. Revisa los nombres de las columnas.", "error"
    except Exception as e:
        return f"Ocurrió un error inesperado: {e}", "error"

# --- Funciones y clase para la GUI ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Resumir Certificados de Excel")
        self.geometry("450x220")
        
        self.ruta_archivo = ""
        self.crear_widgets()

    def crear_widgets(self):
        # Frame principal para organizar los widgets
        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        # Etiqueta para el nombre de la hoja
        lbl_hoja = ttk.Label(frame, text="Nombre de la hoja de cálculo:", anchor="w")
        lbl_hoja.pack(fill=tk.X, pady=(0, 2))

        # Campo de entrada para el nombre de la hoja
        self.entry_hoja = ttk.Entry(frame)
        self.entry_hoja.insert(0, "Hoja1")
        self.entry_hoja.pack(fill=tk.X, pady=(0, 10))

        # Etiqueta para mostrar el archivo seleccionado
        self.lbl_archivo = ttk.Label(frame, text="Por favor, selecciona un archivo de Excel (.xls).", anchor="w")
        self.lbl_archivo.pack(fill=tk.X, pady=(0, 2))

        # Botón para seleccionar el archivo
        btn_seleccionar = ttk.Button(frame, text="Seleccionar archivo", command=self.seleccionar_archivo)
        btn_seleccionar.pack(fill=tk.X, pady=5)
        
        # Botón para procesar el archivo
        self.btn_procesar = ttk.Button(frame, text="Procesar archivo", command=self.iniciar_procesamiento, state=tk.DISABLED)
        self.btn_procesar.pack(fill=tk.X, pady=5)
        
        # Barra de progreso (oculta inicialmente)
        self.progress_bar = ttk.Progressbar(frame, mode='indeterminate')
        
    def seleccionar_archivo(self):
        self.ruta_archivo = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=(("Archivos de Excel", "*.xls"), ("Todos los archivos", "*.*"))
        )
        if self.ruta_archivo:
            self.lbl_archivo.config(text=f"Archivo seleccionado: {self.ruta_archivo}")
            self.btn_procesar.config(state=tk.NORMAL)
        else:
            self.lbl_archivo.config(text="Por favor, selecciona un archivo de Excel (.xls).")
            self.btn_procesar.config(state=tk.DISABLED)

    def iniciar_procesamiento(self):
        if not self.ruta_archivo:
            messagebox.showwarning("Advertencia", "Por favor, primero selecciona un archivo.")
            return

        # Deshabilitar botones, cambiar el texto y mostrar la barra de progreso
        self.btn_procesar.config(text="Procesando...", state=tk.DISABLED)
        self.progress_bar.pack(fill=tk.X, pady=5)
        self.progress_bar.start(10) # El número es el intervalo de actualización en ms
        
        # Ejecutar el procesamiento en un hilo separado
        threading.Thread(target=self.procesar_en_hilo, daemon=True).start()

    def procesar_en_hilo(self):
        nombre_hoja_usuario = self.entry_hoja.get()
        mensaje, tipo = resumir_certificados(self.ruta_archivo, nombre_hoja=nombre_hoja_usuario)
        
        self.after(0, self.finalizar_procesamiento, mensaje, tipo)

    def finalizar_procesamiento(self, mensaje, tipo):
        # Detener la barra de progreso, habilitar botones y mostrar el mensaje
        self.progress_bar.stop()
        self.progress_bar.pack_forget() # Ocultar la barra de progreso
        self.btn_procesar.config(text="Procesar archivo", state=tk.NORMAL)
        
        if tipo == "success":
            messagebox.showinfo("Éxito", mensaje)
        else:
            messagebox.showerror("Error", mensaje)

if __name__ == "__main__":
    app = App()
    app.mainloop()