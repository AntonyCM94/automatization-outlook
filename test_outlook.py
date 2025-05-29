import win32com.client
import logging

# Configuración del registro
logging.basicConfig(
    filename='log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

try: 
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Bandeja de entrada
    
    logging.info("Conexion a Outlook establecida.")
    for item in inbox.Items:
        if item.Class == 43:  # 43 es el tipo de objeto para correos electrónicos
            print(f"De: {item.SenderName}, Asunto: {item.Subject}, Recibido: {item.ReceivedTime}")
            logging.info(f"Correo electrónico encontrado: De: {item.SenderName}, Asunto: {item.Subject}, Recibido: {item.ReceivedTime}")
        elif item.Class == 26:  # 26 es el tipo de objeto para citas
            print(f"Cita: {item.Subject}, Comienza: {item.Start}, Termina: {item.End}")
            logging.info(f"Cita encontrada: Asunto: {item.Subject}, Comienza: {item.Start}, Termina: {item.End}")
        else:
            print("Elemento no es un correo electrónico.")
            logging.info("Elemento ignorado, no es un correo electrónico.")
except Exception as e:
    logging.error(f"Error al conectar a Outlook: {e}")
    logging.exception(e)
# Prueba de conexión a Outlook y listado de correos electrónicos
    print(f"Error al conectar a Outlook: {e}")
# Este código se conecta a Outlook y lista los correos electrónicos en la bandeja de entrada.
# Asegúrate de tener instalado el paquete pywin32 para ejecutar este código
print("Prueba de conexión a Outlook completada.")
print(f"Total de correos electrónicos en la bandeja de entrada: {len(inbox.Items)}")
print("Fin del script de prueba.")