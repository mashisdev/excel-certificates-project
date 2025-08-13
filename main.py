import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time

def resumir_certificados(ruta_archivo, nombre_hoja='Hoja1'):
    """
    Agrupa registros de un archivo de Excel por 'certificado' y consolida los datos.
    Suma los valores únicos de cada monto, considerando la 'secuencia', 
    excepto para 'monto_certificado'.
    
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
        
        # Las columnas que se deben sumar en base a 'certificado' y 'secuencia'
        columnas_monto = ['monto_comp_anual', 'compromiso', 'devengado', 'girado']

        # Creamos una función de agregación personalizada para sumar solo valores únicos
        def sum_unique_per_group(series):
            return series.drop_duplicates().sum()
        
        # Primero, agrupar por 'certificado' y 'secuencia' y sumar los valores únicos
        agregaciones_montos_unicos = {col: sum_unique_per_group for col in columnas_monto}
        df_agrupado_secuencia = df.groupby(['certificado', 'secuencia']).agg(agregaciones_montos_unicos).reset_index()
        
        # Luego, agrupar el resultado por 'certificado' y sumar todos los valores.
        df_final = df_agrupado_secuencia.groupby('certificado')[columnas_monto].sum().reset_index()
        
        # Ahora, manejar las otras columnas que no son montos y la excepción de 'monto_certificado'
        
        # Para las columnas que se mantienen (tomando el primer valor)
        columnas_first = ['ano_eje', 'sec_ejec', 'certificacion', 'expediente', 'fecha_certi', 'ciclo', 'fase']
        df_info = df.groupby('certificado')[columnas_first].first().reset_index()
        
        # Para 'moneda', unir los valores únicos y renombrar la columna
        df_moneda = df.groupby('certificado')['moneda'].apply(lambda x: ', '.join(x.dropna().unique())).reset_index()
        df_moneda.rename(columns={'moneda': 'moneda_agregada'}, inplace=True)
        
        # Para 'monto_certificado', sumar los valores únicos sin importar la secuencia
        def sum_unique(series):
            return series.drop_duplicates().sum()
            
        df_monto_certi = df.groupby('certificado')['monto_certificado'].apply(sum_unique).reset_index()
        
        # Unir todos los DataFrames resultantes
        df_resumido = pd.merge(df_info, df_moneda, on='certificado', how='left')
        df_resumido = pd.merge(df_resumido, df_monto_certi, on='certificado', how='left')
        df_resumido = pd.merge(df_resumido, df_final, on='certificado', how='left')

        # Definir el orden original de las columnas
        orden_original = list(df.columns)
        
        # Se elimina la columna 'secuencia' y se reemplaza 'moneda' por 'moneda_agregada'
        orden_final = [col for col in orden_original if col != 'secuencia']
        
        # Reemplazar el nombre de la columna en la lista de orden final
        if 'moneda' in orden_final:
            orden_final[orden_final.index('moneda')] = 'moneda_agregada'
        
        df_resumido = df_resumido[orden_final]
        
        # Renombrar 'moneda_agregada' a 'moneda' para el archivo de salida
        df_resumido.rename(columns={'moneda_agregada': 'moneda'}, inplace=True)
        
        df_resumido = df_resumido.sort_values(by='certificado')
        
        nombre_archivo_salida = 'certificados_resumidos.xlsx'
        df_resumido.to_excel(nombre_archivo_salida, index=False)
        
        return f"Proceso completado. Los datos consolidados se han guardado en '{nombre_archivo_salida}'.", "success"
        
    except FileNotFoundError:
        return "Error: El archivo no se encontró. Asegúrate de que el archivo exista.", "error"
    except KeyError as e:
        # Aquí se incluye la variable e para que muestre el nombre de la columna que falta
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