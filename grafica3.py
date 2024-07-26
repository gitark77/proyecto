import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def load_data(file):
    try:
        df = pd.read_excel(file)
        st.write("Datos cargados exitosamente.")
        return df
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return None

def col_letter_to_index(letter):
    """Convierte una letra o secuencia de letras en un índice de columna (0-indexado)."""
    letter = letter.upper()
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def get_range(df):
    # Seleccionar rango de filas
    start_row = st.number_input("Desde qué fila tomar los datos (empieza desde 1)", min_value=1, value=1)
    end_row = st.number_input("Hasta qué fila tomar los datos (empieza desde 1)", min_value=1, value=df.shape[0])

    if start_row and end_row:
        start_row -= 1  # Convertir a índice basado en 0
        end_row -= 1
        df = df.iloc[start_row:end_row + 1]
        
        # Seleccionar rango de columnas
        start_col = st.text_input("Desde qué columna tomar los datos (letra, e.g., 'A')")
        end_col = st.text_input("Hasta qué columna tomar los datos (letra, e.g., 'Z')")

        if start_col and end_col:
            start_idx = col_letter_to_index(start_col)
            end_idx = col_letter_to_index(end_col)
            df = df.iloc[:, start_idx:end_idx + 1]
            st.write("Rango de datos ajustado.")
    
    return df

def calculate_statistics(df):
    if df is not None:
        # Calcular promedio de cada columna
        averages = df.mean(numeric_only=True)
        st.write(f"Promedios calculados: {averages}")
        return averages
    return None

def plot_data(df, averages):
    if df is not None and averages is not None:
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

        st.pyplot(fig)

def main():
    st.title("Visualización de Datos desde Excel")
    
    uploaded_file = st.file_uploader("Elige un archivo Excel", type="xlsx")
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        if df is not None:
            df = get_range(df)
            averages = calculate_statistics(df)
            plot_data(df, averages)

if __name__ == "__main__":
    main()
