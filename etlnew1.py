import pandas as pd
import os
import re
from tkinter import Tk, filedialog, simpledialog, messagebox, ttk
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from tkinter.constants import N, W, E, S

def seleccionar_carpeta():
    root = Tk()
    root.withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con archivos Excel")
    root.destroy()  
    return carpeta

def obtener_parametros():
    root = Tk()
    root.withdraw()
    columnas = simpledialog.askstring("Input", "Ingresa el rango de columnas (A:Z)")
    fila_inicial = simpledialog.askinteger("Input", "Ingresa el número de la fila a iniciar")
    root.destroy() 
    return columnas, fila_inicial

def extraer_fecha(nombre_archivo):
    match = re.search(r'\d{4}\.\d{2}\.\d{2}', nombre_archivo)
    if match:
        fecha = match.group().split('.')
        return int(fecha[0]), int(fecha[1]), int(fecha[2])
    return None, None, None

def cargar_datos(carpeta, columnas, fila_inicial):
    all_files = [os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    df_list = []
    
    def process_file(file):
        try:
            anio, mes, dia = extraer_fecha(file)
            df = pd.read_excel(file, sheet_name='ITEM_O', usecols=columnas, skiprows=fila_inicial-1)
            df['AÑO'] = anio
            df['MES'] = mes
            df['DIA'] = dia
            return df
        except FileNotFoundError:
            print(f"Archivo no encontrado en la ruta: {file}")
        except ValueError as ve:
            print(f"Error al leer el archivo Excel {file}: {ve}")
        except Exception as e:
            print(f"Error al procesar el archivo {file}: {e}")
        return pd.DataFrame()

    with ThreadPoolExecutor() as executor:
        results = list(tqdm(executor.map(process_file, all_files), total=len(all_files), desc="Procesando archivos excel"))
        df_list.extend(results)
    
    if df_list:
        df_consolidado = pd.concat(df_list, ignore_index=True)
    else:
        df_consolidado = pd.DataFrame()
    
    return df_consolidado

def mostrar_dataframe(df):
    root = Tk()
    root.title("DataFrame Consolidado")

    frame = ttk.Frame(root, padding="3 3 12 12")
    frame.grid(column=0, row=0, sticky=(N, W, E, S))

    tree = ttk.Treeview(frame, columns=list(df.columns), show='headings')

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for idx, row in df.iterrows():
        tree.insert("", "end", values=list(row))

    tree.grid(column=0, row=0, sticky=(N, W, E, S))

    save_button = ttk.Button(frame, text="Guardar", command=lambda: guardar_dataframe(df))
    save_button.grid(column=0, row=1, sticky=(W, E))

    root.mainloop()

def guardar_dataframe(df):
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    
    if file_path:
        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", f"El archivo {os.path.basename(file_path)} se ha guardado exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")
    else:
        messagebox.showerror("Error", "No se seleccionó ninguna ubicación para guardar el archivo.")

def main():
    carpeta = seleccionar_carpeta()
    if not carpeta:
        messagebox.showerror("Error", "Debes seleccionar una carpeta con archivos Excel.")
        return
    
    columnas, fila_inicial = obtener_parametros()
    if not columnas or not fila_inicial:
        messagebox.showerror("Error", "No se proporcionaron todos los parámetros necesarios para procesar.")
        return
    
    df_consolidado = cargar_datos(carpeta, columnas, fila_inicial)
    
    if df_consolidado.empty:
        messagebox.showerror("Error", "No se encontraron datos en los archivos proporcionados para procesar.")
        return
    
    mostrar_dataframe(df_consolidado)

if __name__ == "__main__":
    main()
