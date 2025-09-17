import pandas as pd
data = pd.read_csv('control_emails.csv')

# # # Eliminar las columnas 4 y 5
data = data.iloc[:, [0,1,2,3]]

data.to_csv('control_emails.csv', index=False)

# import win32com.client

# Lectura de archivo .csv con datos de control de correos
# df = pd.read_csv('control_emails.csv')
# sistemas = df.iloc[:, 0].tolist()

# # Conectar a Outlook
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# nombre_cuenta = "energias.renovables.es@dekra.com"
# store = outlook.Stores[nombre_cuenta]
# bandeja_entrada = store.GetDefaultFolder(6) # Acceso a la bandeja de entrada
# carpeta_dades_meteo = bandeja_entrada.Folders["Dades Meteo"] # Acceso a la subcarpeta "Dades Meteo"
# mensajes = carpeta_dades_meteo.Items # Items de la carpeta
# mensajes.Sort("[ReceivedTime]", True)

# for msg in mensajes:
#     print(msg.SenderEmailAddress)
#     remitente = msg.SenderEmailAddress
#     if remitente == "status@support.zxlidars.com":
#            print("Es un correo de ZX300 con asunto ", msg.Subject)
#     else:
#         continue