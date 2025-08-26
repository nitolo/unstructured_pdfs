from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter
import pytesseract
import re
import pandas as pd
import fitz
import os
import requests
import sys
import ollama
import requests
import json
import time
print(sys.executable)

def ensure_model_ready():
   """Pre-carga y verifica que el modelo esté en memoria"""
   print("Verificando modelo...")
   start = time.time()
   
   # Prompt mínimo para despertar el modelo
   payload_test = {
       #"model": "llama3.2:1b",
       "model": "llama3.2:3b",
       #"prompt": "Responde solo: OK",
       "prompt": "Extrae datos del siguiente texto en JSON: tasa_fwd: 4000, valor_nominal_usd: 1000000, fecha_inicio: 01/01/2024",
       "stream": False,
       "options": {
           "num_predict": 100,
           "temperature": 0.1,
           "keep_alive": "30m"
       }
   }
   
   try:
       response = requests.post(
           "http://localhost:11434/api/generate",
           json=payload_test,
           timeout=30
       )
       result = response.json()
       respuesta_texto = result.get('response', '')
       print(respuesta_texto)
       elapsed = time.time() - start
       
       if elapsed > 10:
           print(f"Modelo se cargó desde disco ({elapsed:.1f}s)")
       else:
           print(f"Modelo ya estaba en memoria ({elapsed:.1f}s)")
           
   except Exception as e:
       print(f"Error pre-cargando modelo: {e}")

# PRE-CARGAR MODELO AL INICIO
ensure_model_ready()

# Cuando esten replicando el code, no se les olvide colcoar la ruta donde esta su tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\ntorreslo\AppData\Local\Programs\Tesseract-OCR\tesseract'

# Hay que instalar Poppler, too
directory_poppler = r'C:\Poppler\poppler-24.08.0\Library\bin'

# Voy a definir un PDF super feo para reconocer
#file_path = r'Z:\17. Reporting Automation\Cartas NDFs\BANCO SANTANDER 17-06-25 1010.pdf'
#file_path = r'Z:\17. Reporting Automation\Cartas NDFs\Cartas sin firmas\2025\07\280725\Confirmation-AE2025071091434370.pdf'
file_path = r"Z:\17. Reporting Automation\Cartas NDFs\Cartas sin firmas\2025\08\110825\2025-07-24_COLOMBIA TELECO.pdf"
filename = os.path.basename(file_path)
filename

if "Confirmation-AE" in filename: # jpmorgan
  doc = fitz.open(file_path)
  text = doc.load_page(1).get_text()
  doc.close()
  banco = "JPMORGAN"
  idioma = 'english'
  print(text)
  print(banco)
elif "COLOMBIA TELECO" in filename: # bancolombia
  reader = PdfReader(file_path)
  writer = PdfWriter()
  if reader.is_encrypted:
      reader.decrypt("830122566")
  # Add all pages to the writer
  for page in reader.pages:
      writer.add_page(page)
  # Save the new PDF to a file
  with open(file_path, "wb") as f:
      writer.write(f)
  # Volver a abrir el archivo 
  doc = fitz.open(file_path)
  text = doc.load_page(0).get_text()
  doc.close()
  banco = "BANCOLOMBIA"
  idioma = 'spanish'
  print(text)
  print(banco)
# here, ya empieza la parte de los pdfs escaneados
elif re.search(r"\b\d{7}\b", filename): # scotiabank
  images = convert_from_path(file_path, dpi=400, poppler_path=directory_poppler)
  text= ""
  for image in images:
    text += pytesseract.image_to_string(image, lang = 'eng+spa')
  banco = "SCOTIABANK"
  idioma = 'spanish'
  print(text)
  print(banco)
elif "_NDFV_FW" in filename: # ITAU
  images = convert_from_path(file_path, dpi=500, poppler_path=directory_poppler)
  text= ""
  for image in images:
    text += pytesseract.image_to_string(image, lang = 'eng+spa')
  banco = "ITAÚ"
  idioma = 'spanish'
  print(text)
  print(banco)
else:
  print("O falta introducir el banco, o el PDF ya está con el nombre correcto")

# Tener la lista de matrones para reconocer los bancos
bank_patterns = [
    r"(SCOTIABANK)",
    r"(DAVIVIENDA)",
    r"(ITAÚ)",
    r"(BANCOLOMBIA)",
    r"(JPMorgan)",
    r"(BANCO DE OCCIDENTE)",
    r"(BANCO SANTANDER)",
    r"(CITIBANK COLOMBIA)"
]

# Acá es reconocer el pattern del banco
for pattern in bank_patterns:
    match = re.search(pattern, text)
    if match:
        print(match.group(1).strip().upper())

##########################
# LIMPIEZA DE TEXTO
##########################
# llama 3.2b ya se comporta bien, pero yo necesito que se me demore menos

#text = ' '.join(words) if isinstance(text, list) else text
# quitar ruido
text_clean = re.sub(r'[^a-zA-Z0-9\s,.:-]', '', text)
# normalizar saltos de linea
text_clean = re.sub(r'[^\w\s,.:-]', '', text_clean)
# elimina caracteres no ASCII
#text_clean= re.sub(r'[^\x00-\x7F]+', '', text_clean)  
#print(text_clean)

# elimina saltos de línea y tabulaciones
#text_clean = re.sub(r'\n|\t', ' ', text_clean)
#print(text_clean)  

#print(text)
#print(text_clean)

# Here, ya coloco el prompt para poder extraer. 
# Siempre es bueno estar mirandolo.
prompt = f"""Eres un experto en análisis de contratos financieros en inglés y español con experiencia en extracción de datos estructurados.

OBJETIVO: Extraer información específica de contratos forward y presentarla en formato JSON.

CAMPOS A EXTRAER:
1. tasa_fwd: Tasa Forward (número decimal)
2. valor_nominal_usd: Valor nominal en USD (número entero)  
3. fecha_inicio: Fecha de inicio/negociación (formato dd/mm/aaaa)

DEFINICIONES ESPECÍFICAS:

tasa_fwd:
- Es la tasa de cambio forward/strike del contrato
- NO confundir con el valor total en COP
- Buscar valores típicos entre 3000-5000 para COP/USD
- Puede aparecer como "Tasa", "Strike", "Rate", "Forward Rate"

valor_nominal_usd:
- Es el monto nocional/principal en dólares estadounidenses
- SIEMPRE acompañado de "USD" o indicado en columna de moneda USD
- NO es el equivalente en COP (que será mucho mayor)
- Buscar valores como "1,000,000.00 USD", "2,500,000 USD"
- IGNORAR valores en COP (que son resultado de multiplicar tasa x nominal)

fecha_inicio:
- Fecha de negociación o trade date
- Puede aparecer como "Trade Date", "Fecha Negociación", "Deal Date"

REGLAS DE FORMATO CRÍTICAS:
- tasa_fwd: ELIMINAR puntos de miles, mantener coma decimal
  Correcto: 4236,20
  Incorrecto: 4.236,20
  
- valor_nominal_usd: SOLO números enteros, sin separadores
  Correcto: 2000000
  Incorrecto: 2,000,000 o 2.000.000
  
- fecha_inicio: Formato estricto dd/mm/aaaa
  Correcto: 15/03/2024
  Incorrecto: 15-03-2024 o 2024/03/15 o 15 de mayo de 2024 o March 15th 2024 o March 15, 2024 o 15-Mar-2024

EJEMPLOS DE TRANSFORMACIÓN:
- "4.236,20" → 4236,20
- "2,000,000.00 USD" → 2000000   
- "March 15, 2024" → 15/03/2024

INSTRUCCIONES DE SALIDA:
- Responde ÚNICAMENTE con JSON válido
- No incluyas explicaciones, comentarios o texto adicional
- Si un campo no se encuentra, usa null
- Mantén exactamente estos nombres de campo

FORMATO DE RESPUESTA ESPERADO:
{{
  "tasa_fwd": 4236.20,
  "valor_nominal_usd": 2000000,
  "fecha_inicio": "15/03/2024"
}}

TEXTO DEL CONTRATO A ANALIZAR:
---
{text_clean}
---
Procede con la extracción:"""

# Configuración para Ollama
 
print("=== CONFIGURACIÓN ORIGINAL ===")
start_time = time.time()
payload = {
    #"model": "llama3.2:1b",
    "model": "llama3.2:3b",
    "prompt": prompt,
    "stream": False,
    "options": {
        "temperature": 0.1,
        "keep_alive": "30m"
    }
}

print("Enviando petición a Ollama...")

# Hacer la petición
response = requests.post(
    "http://localhost:11434/api/generate",
    json=payload,
    timeout=180)

result = response.json()

respuesta_texto = result.get('response', '')

print(respuesta_texto)
end_time = time.time()

tiempo_original = end_time - start_time
print(f"Tiempo original: {tiempo_original:.2f} segundos")