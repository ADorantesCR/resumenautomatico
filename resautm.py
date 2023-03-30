!pip install sumy
!pip install openpyxl

import pandas as pd
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from sumy.utils import get_stop_words
from openpyxl import Workbook

# Función para generar el resumen de una nota
def generate_summary(text):
    # Crear un objeto parser usando el texto de entrada
    parser = PlaintextParser.from_string(text, Tokenizer("spanish"))
    # Crear un objeto summarizer
    summarizer = LsaSummarizer()
    # Obtener las palabras de parada para español
    stop_words = get_stop_words("spanish")
    # Añadir las palabras de parada al summarizer
    summarizer.stop_words = stop_words
    # Generar un resumen de 3 oraciones
    summary = summarizer(parser.document, 3)
    # Unir las oraciones del resumen en un string
    summary = " ".join(str(sentence) for sentence in summary)
    return summary

# Cargar el archivo csv con pandas
df = pd.read_csv('notas.csv')

# Generar un resumen para cada nota en la columna 'notas' del dataframe
summaries = []
for note in df['notas']:
    summary = generate_summary(note)
    summaries.append(summary)

# Agregar los resúmenes al dataframe
df['resumen'] = summaries

# Guardar el dataframe en un archivo de Excel
with pd.ExcelWriter('resumenes.xlsx') as writer:
    df.to_excel(writer, sheet_name='Resúmenes', index=False)
