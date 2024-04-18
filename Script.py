import os
from docx import Document
import pandas as pd
import re

def limpiar_texto(texto):
    # Reemplaza saltos de línea, tabulaciones y múltiples espacios por un espacio simple
    texto = re.sub(r'[\t\n\r]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

path = 'windows-directory'
files = [os.path.join(path, f) for f in os.listdir(path) if f.endswith(".docx")]
texts = []
file_names = []

for file in files:
    doc = Document(file)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(limpiar_texto(para.text))
    condensedText = ' '.join(fullText)
    texts.append(condensedText)
    
    # Obtén el nombre base del archivo, quita la extensión .docx y agrégalo a la lista
    base_name = os.path.splitext(os.path.basename(file))[0]  # [0] nos da el nombre sin la extensión
    file_names.append(base_name)

df = pd.DataFrame({'texto': texts, 'Fuente': file_names})
df.to_excel('xlsx_files', index=False)
