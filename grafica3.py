import tkinter as tk
from tkinter import filedialog, simpledialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

def load_data():
    global df
    file_path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        sheet_name = simpledialog.askstring("Input", "Nombre de la hoja de cálculo (deja vacío para usar la primera hoja)")
        if not sheet_name:
            sheet_name = 0  # Usar la primera hoja si no se especifica ningún nombre
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print("Datos cargados exitosamente.")
        except Exception as e:
            print(f"Error al cargar datos: {e}")

def col_letter_to_index(letter):
    """Convierte una letra o secuencia de letras en un índice de columna (0-indexado)."""
    letter = letter.upper()
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def get_range():
    global df
    if df is not None:
        # Seleccionar rango de filas
        start_row = simpledialog.askinteger("Input", "Desde qué fila tomar los datos (empieza desde 1)")
        end_row = simpledialog.askinteger("Input", "Hasta qué fila tomar los datos (empieza desde 1)")

        if start_row is not None and end_row is not None:
            start_row -= 1  # Convertir a índice basado en 0
            end_row -= 1
            df = df.iloc[start_row:end_row + 1]
        
        # Seleccionar rango de columnas
        start_col = simpledialog.askstring("Input", "Desde qué columna tomar los datos (letra, e.g., 'A')")
        end_col = simpledialog.askstring("Input", "Hasta qué columna tomar los datos (letra, e.g., 'Z')")

        if start_col and end_col:
            start_idx = col_letter_to_index(start_col)
            end_idx = col_letter_to_index(end_col)
            df = df.iloc[:, start_idx:end_idx + 1]
            print("Rango de datos ajustado.")

def calculate_statistics():
    global df
    if df is not None:
        # Calcular promedio de cada columna
        averages = df.mean(numeric_only=True)
        print(f"Promedios calculados: {averages}")
        return averages
    return None

def plot_data(averages):
    global df
    if df is not None and averages is not None:
        # Crear la ventana principal
        root = tk.Tk()
        root.title("Visualización de Datos")

        # Crear un marco para contener los gráficos
        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True)

        # Crear los gráficos
        fig, axs = plt.subplots(1, 3, figsize=(18, 6))

        # Graficar promedio de cada columna
        averages.plot(kind='bar', ax=axs[0])
        axs[0].set_title('Promedio de cada columna')
        axs[0].set_xlabel('Columnas')
        axs[0].set_ylabel('Promedio')

        # Graficar datos mayores
        numeric_data = df.select_dtypes(include=[np.number])
        max_values = numeric_data.max()
        max_values.plot(kind='bar', color='orange', ax=axs[1])
        axs[1].set_title('Datos mayores de cada columna')
        axs[1].set_xlabel('Columnas')
        axs[1].set_ylabel('Valor Máximo')

        # Graficar datos en gráfico de torta
        if not max_values.empty:
            axs[2].pie(max_values, labels=max_values.index, autopct='%1.1f%%', startangle=140)
            axs[2].set_title('Distribución de valores mayores')

        # Integrar matplotlib con tkinter
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Función para guardar los gráficos
        def save_graphs():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("JPG files", "*.jpg"), ("All files", "*.*")]
            )
            if file_path:
                fig.savefig(file_path)
                print(f"Gráficos guardados en {file_path}")

        # Botón para guardar los gráficos
        save_button = tk.Button(root, text="Guardar Gráficos", command=save_graphs)
        save_button.pack(pady=10)

        # Iniciar el bucle principal de tkinter
        root.mainloop()

def main():
    global df
    df = None

    load_data()
    get_range()
    averages = calculate_statistics()
    plot_data(averages)

if __name__ == "__main__":
    main()
