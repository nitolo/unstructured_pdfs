import os
import win32com.client as win32
from datetime import datetime
from pathlib import Path

# Fecha actual
today = datetime.now()
year = today.strftime("%Y")
month = today.strftime("%m")
day = today.strftime("%d%m%y")
fecha_hoy = today.strftime("%d/%m/%Y")

# Ruta base
base_dir = Path(r"Z:\17. Reporting Automation\Cartas NDFs\Cartas para firmar")
output_dir = base_dir / year / month / day

# Inicializar Outlook
outlook = win32.Dispatch('outlook.application')
namespace = outlook.GetNamespace("MAPI")


# Recorrer subcarpetas
for subfolder in output_dir.iterdir():
    if subfolder.is_dir():
        pdf_files = list(subfolder.glob("*.pdf"))
        if pdf_files:
            # Crear correo
            mail = outlook.CreateItem(0)
            mail.Subject = f"Cartas NDF - {subfolder.name} {fecha_hoy}"
            mail.To = "zamir.suz@telefonica.com"
            cc = ["gerson.diaz@telefonica.com", "wilson.garavito@telefonica.com",
                  "daniel.siachoque@telefonica.com", "tania.manzano@telefonica.com"
                  , "mercadodecapitales.co@telefonica.com", "n.torres01@telefonica.com"]
            mail.CC = "; ".join(cc)
            mail.SentOnBehalfOfName = "n.torres01@telefonica.com"
            mail.Body = f"Hola Zamir,\n\nAdjunto confirmaciones NDFs correspondientes a {subfolder.name} para su firma.\n\nMil gracias,\nEquipo de Mercado de Capitales"
            #mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Establecer cuenta de envío

            # Adjuntar PDFs
            for pdf in pdf_files:
                mail.Attachments.Add(str(pdf))

            # Enviar correo
            mail.Send()

print("Correos enviados correctamente.")

