import os
import re
import json
import shutil
import time
import requests
import pandas as pd
import pytesseract
import fitz
from datetime import datetime
from pathlib import Path
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter

# Configuraciones
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\ntorreslo\AppData\Local\Programs\Tesseract-OCR\tesseract'
directory_poppler = r'C:\Poppler\poppler-24.08.0\Library\bin'

# Fechas y directorios
today = datetime.now()
year = today.strftime("%Y")
month = today.strftime("%m")
day = today.strftime("%d%m%y")
fecha_proceso = today.strftime('%d-%m-%y')

input_dir = Path(r"Z:/17. Reporting Automation/Cartas NDFs/Cartas sin firmas") / year / month / day
output_base_dir = Path(r"Z:/17. Reporting Automation/Cartas NDFs/Cartas para firmar")
excel_output_dir = Path(r"Z:/17. Reporting Automation/Cartas NDFs/Cartas para firmar")
excel_output_dir.mkdir(parents=True, exist_ok=True)

def ensure_model_ready():
    """Pre-carga el modelo Ollama"""
    print("Verificando modelo...")
    payload_test = {
        "model": "llama3.2:3b",
        "prompt": "Extrae datos del siguiente texto en JSON: tasa_fwd: 4000, valor_nominal_usd: 1000000, fecha_inicio: 01/01/2024",
        "stream": False,
        "options": {
            "num_predict": 100,
            "temperature": 0.1,
            "keep_alive": "30m"
        }
    }
    
    try:
        response = requests.post("http://localhost:11434/api/generate", json=payload_test, timeout=30)
        result = response.json()
        print("Modelo listo:", result.get('response', '')[:50])
    except Exception as e:
        print(f"Error pre-cargando modelo: {e}")

def extract_text_from_pdf(file_path):
    """Extrae texto según el tipo de PDF"""
    filename = os.path.basename(file_path)
    
    if "Confirmation-AE" in filename:  # JPMorgan
        doc = fitz.open(file_path)
        text = doc.load_page(1).get_text()
        doc.close()
        banco = "JPMORGAN"
        
    elif "COLOMBIA TELECO" in filename:  # Bancolombia
        reader = PdfReader(file_path)
        writer = PdfWriter()
        if reader.is_encrypted:
            reader.decrypt("830122566")
        for page in reader.pages:
            writer.add_page(page)
        with open(file_path, "wb") as f:
            writer.write(f)
        doc = fitz.open(file_path)
        text = doc.load_page(0).get_text()
        doc.close()
        banco = "BANCOLOMBIA"
        
    elif re.search(r"\b\d{7}\b", filename):  # Scotiabank
        images = convert_from_path(file_path, dpi=600, poppler_path=directory_poppler)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image, lang='eng+spa')
        banco = "DAVIbank"
        
    elif "_NDFV_FW" in filename:  # ITAÚ
        images = convert_from_path(file_path, dpi=500, poppler_path=directory_poppler)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image, lang='eng+spa')
        banco = "ITAÚ"
        
    else:  # Otros bancos con OCR
        images = convert_from_path(file_path, dpi=500, poppler_path=directory_poppler)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image, lang='eng+spa')
        banco = detect_banco_from_text(text)
    
    return text, banco

def detect_banco_from_text(text):
    """Detecta el banco del texto extraído"""
    bank_patterns = [
        r"(SCOTIABANK)", r"(DAVIVIENDA)", r"(ITAÚ)", r"(BANCOLOMBIA)",
        r"(JPMorgan)", r"(BANCO DE OCCIDENTE)", r"(BANCO SANTANDER)", 
        r"(CITIBANK COLOMBIA)", r"(CORFICOLOMBIANA)", r"(DAVIbank)"
    ]
    
    for pattern in bank_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip().upper()
    return "BANCO_DESCONOCIDO"

def clean_text(text):
    """Limpia el texto para mejor procesamiento"""
    text_clean = re.sub(r'[^a-zA-Z0-9\s,.:-]', '', text)
    text_clean = re.sub(r'[^\w\s,.:-]', '', text_clean)
    return text_clean

def extract_with_llm(text):
    """Extrae datos usando Ollama"""
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
- Buscar valores típicos entre 3000-5000 para COP/USD. Puede existir el caso que sea USD/EUR.
- Puede aparecer como "Tasa", "Strike", "Rate", "Forward Rate", "Tasa Forward"

valor_nominal_usd:
- Es el monto nocional/principal/valor negociado en dólares estadounidenses.
- SIEMPRE acompañado de "USD" o indicado en columna de moneda USD
- NO es el equivalente en COP (que será mucho mayor)
- Buscar valores como "1,000,000.00 USD", "2,500,000 USD", "1,000,000.00"
- IGNORAR valores en COP (que son resultado de multiplicar tasa x nominal)

fecha_inicio:
- Fecha de negociación o trade date
- Puede aparecer como "Trade Date", "Fecha Negociación", "Deal Date"

REGLAS DE FORMATO CRÍTICAS:
- tasa_fwd: ELIMINAR puntos de miles, mantener coma decimal
  Correcto: 4236,20
  Incorrecto: 4.236,20 o 4.233620
  
- valor_nominal_usd: SOLO números enteros, sin separadores
  Correcto: 2000000
  Incorrecto: 2,000,000 o 2.000.000
  
- fecha_inicio: Formato estricto dd/mm/aaaa
  Correcto: 15/03/2024
  Incorrecto: 15-03-2024 o 2024/03/15 o 05082025 o 15 de mayo de 2024

INSTRUCCIONES DE SALIDA:
- Responde ÚNICAMENTE con JSON válido
- No incluyas explicaciones, comentarios o texto adicional
- Si un campo no se encuentra, usa null

FORMATO DE RESPUESTA ESPERADO:
{{
  "tasa_fwd": 4236.20,
  "valor_nominal_usd": 2000000,
  "fecha_inicio": "15/03/2024"
}}

TEXTO DEL CONTRATO A ANALIZAR:
---
{text[:3000]}
---
Procede con la extracción:"""

    payload = {
        "model": "llama3.2:3b",
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.1,
            "keep_alive": "30m"
        }
    }

    try:
        response = requests.post("http://localhost:11434/api/generate"
        , json=payload
        , timeout=280 # importantisimo acomodar esto. muchas veces toca cambiarlo por la potencia del compu
        )
        result = response.json()
        respuesta_texto = result.get('response', '')
        
        # Buscar JSON en la respuesta
        json_match = re.search(r'\{.*\}', respuesta_texto, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(0))
        else:
            return {"error": "No se encontró JSON válido"}
            
    except Exception as e:
        return {"error": f"Error con Ollama: {str(e)}"}

def main():
    # Verificar modelo
    ensure_model_ready()
    
    # Lista para almacenar resultados
    records = []
    current_id = 1001
    
    print(f"Procesando PDFs en: {input_dir}")
    
    # Verificar si existe el directorio
    if not input_dir.exists():
        print(f"Error: El directorio {input_dir} no existe")
        return
    
    pdf_files = list(input_dir.glob("*.pdf"))
    print(f"Encontrados {len(pdf_files)} archivos PDF")
    
    for pdf_file in pdf_files:
        print(f"\nProcesando: {pdf_file.name}")
        
        # Extraer texto y banco
        text, banco = extract_text_from_pdf(str(pdf_file))
        text_clean = clean_text(text)
        
        # Extraer datos con LLM
        resultado_llm = extract_with_llm(text_clean)
        
        # Crear nuevo nombre y mover archivo
        nuevo_nombre = f"{banco} {fecha_proceso} {current_id}.pdf"
        destino_dir = output_base_dir / year / month / day / banco
        destino_dir.mkdir(parents=True, exist_ok=True)
        nuevo_path = destino_dir / nuevo_nombre
        
        try:
            shutil.copy(str(pdf_file), str(nuevo_path))
            print(f"Movido a: {nuevo_path}")
        except Exception as e:
            print(f"Error moviendo archivo: {e}")
        
        # Crear registro
        record = {
            "id": current_id,
            "archivo_original": pdf_file.name,
            "nuevo_nombre_archivo": nuevo_nombre,
            "banco": banco,
            "fecha_proceso": fecha_proceso
        }
        
        if 'error' in resultado_llm:
            record.update({
                "tasa_fwd": "ERROR",
                "valor_nominal_usd": "ERROR", 
                "fecha_inicio": "ERROR",
                "error": resultado_llm.get("error", "Error desconocido")
            })
        else:
            record.update({
                "tasa_fwd": resultado_llm.get("tasa_fwd", ""),
                "valor_nominal_usd": resultado_llm.get("valor_nominal_usd", ""),
                "fecha_inicio": resultado_llm.get("fecha_inicio", ""),
                "error": ""
            })
        
        records.append(record)
        current_id += 1
        
        print(f"Datos extraídos: {resultado_llm}")
    
    # Guardar Excel
    if records:
        df = pd.DataFrame(records)
        timestamp = datetime.now().strftime('%Y%m%d')
        excel_path = excel_output_dir / year / month / day / f"Seguimiento NDFs {timestamp}.xlsx"
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"\nResultados guardados en: {excel_path}")
    else:
        print("No se procesaron archivos")

if __name__ == "__main__":
    main()


