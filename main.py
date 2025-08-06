import pandas as pd

def resumir_certificados(ruta_archivo, nombre_hoja='ReporteMeta'):
    """
    Agrupa registros de un archivo de Excel por 'certificado' y consolida los datos.

    Args:
        ruta_archivo (str): La ruta al archivo de Excel a procesar.
        nombre_hoja (str): El nombre de la hoja de cálculo.
    """
    try:
        # Usamos dtype para asegurarnos de que estas columnas se lean como texto
        dtype_cols = {
            'sec_ejec': str,
            'certificado': str,
            'certificacion': str,
            'secuencia': str,
            'expediente': str
        }
        df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, dtype=dtype_cols)
        print("Archivo de Excel cargado exitosamente.")

        # Limpiar espacios en blanco de los nombres de las columnas si los hay
        df.columns = df.columns.str.strip()

        # Definir las operaciones de agregación para cada columna
        agregaciones = {
            'ano_eje': 'first',
            'sec_ejec': 'first',
            'certificacion': 'first',
            'secuencia': 'first',
            'ciclo': 'first',
            'fase': 'first',
            'moneda': lambda x: ', '.join(x.dropna().unique()),
            'expediente': 'first',
            'monto_certificado': 'last',
            'monto_comp_anual': 'last',
            'compromiso': 'last',
            'devengado': 'last',
            'girado': 'last',
        }

        # Agrupar el DataFrame por la columna 'certificado' y aplicar las agregaciones
        df_resumido = df.groupby('certificado', as_index=False).agg(agregaciones)

        # Opcional: ordenar el DataFrame final por la columna 'certificado'
        df_resumido = df_resumido.sort_values(by='certificado')

        # Guardar el DataFrame resultante en un nuevo archivo de Excel
        nombre_archivo_salida = 'certificados_resumidos.xlsx'
        df_resumido.to_excel(nombre_archivo_salida, index=False)

        print(f"Proceso completado. Los datos consolidados se han guardado en '{nombre_archivo_salida}'.")

    except FileNotFoundError:
        print(f"Error: El archivo '{ruta_archivo}' no se encontró en la misma carpeta. Por favor, verifica el nombre.")
    except KeyError as e:
        print(f"Error: La columna {e} no se encontró en el archivo. Revisa si los nombres de las columnas son correctos.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

# --- Uso del script ---
if __name__ == "__main__":
    nombre_de_tu_archivo = 'rptCertificacionCompromisoExpediente.xls'
    resumir_certificados(nombre_de_tu_archivo)