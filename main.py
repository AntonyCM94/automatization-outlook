import win32com.client
import datetime

# configuracion 
KEYWORD = "propuesta"
DAYS_BEFORE_FOLLOWUP = 3

# Inicializar outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
sent_folder = outlook.GetDefaultFolder(5)  # 5 es el índice para la carpeta "Elementos enviados"
inbox_folder = outlook.GetDefaultFolder(6)  # 6 es el índice para la carpeta "Bandeja de entrada"

# Fecha limite para seguimiento
cutoff_date = datetime.datetime.now() - datetime.timedelta(days=DAYS_BEFORE_FOLLOWUP)

# Escanear correos enviados
print("Escaneando correos enviados...")
for sent_mail in sent_folder.Items:
    if sent_mail.Class != 43: # 43 es el tipo de objeto para correos electrónicos y asegura que es MailItem
        continue

    if KEYWORD.lower() in sent_mail.Body.lower():
        sent_time = sent_mail.SentOn

        if sent_time < cutoff_date:
            subject = sent_mail.Subject
            recipient = sent_mail.To

            # Buscar respuesta en la bandeja de entrada
            replied = False
            for mail in inbox_folder.Items:
                if mail.Class != 43:
                    continue
                if subject in mail.Subject and recipient in mail.SenderEmailAddress:
                    replied = True
                    break
                if not replied: 
                    print(f"No se ha recibido respuesta para el correo enviado a {recipient} con asunto '{subject}'.")
                    print("Considerar enviar un recordatorio o seguimiento.")
