# import pandas as pd

# data = pd.read_excel('control_emails_padre.xlsx')
# data = data[['Sistema', 'Identificador', 'Remitente']]
# data.to_csv('control_emails.csv', index=False)

# print(data)

import pandas as pd 
import win32com.client
import re 
import os

# Función para extraer la fecha del asunto
def extraer_fecha(asunto: str, patron, remitente: str) -> str | None:
    m = patron.search(asunto)
    if not m:
        return None
    
    if remitente in ["estacionesmeteo", "emailrelay"]:
        return m.group(1)  # ya viene YYYY-MM-DD
    
    elif remitente == "windcube":
        return m.group(1).replace("/", "-")  # YYYY/MM/DD -> YYYY-MM-DD
    
    elif remitente == "zx":
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    
    return None

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
nombre_cuenta = "energias.renovables.es@dekra.com"
store = outlook.Stores[nombre_cuenta]
bandeja_entrada = store.GetDefaultFolder(6)
carpeta_dades_meteo = bandeja_entrada.Folders["Dades Meteo"]
mensajes = carpeta_dades_meteo.Items
mensajes.Sort("[ReceivedTime]", True)

primer_mensaje = mensajes.GetFirst()
#print(primer_mensaje.ReceivedTime)

fecha_objetivo = pd.Timestamp.now()
#fecha_objetivo = pd.Timestamp(2025,8,27,8,0,0) # Formato 
print("Fecha objetivo: ", fecha_objetivo.date())
#fecha_str = fecha_objetivo.strftime("%Y-%m-%d")

#fecha_actual = fecha_objetivo.normalize()  # usar fecha objetivo como referencia
fecha_fin = fecha_objetivo.strftime('%m/%d/%Y %H:%M %p')
fecha_ini = (fecha_objetivo - pd.Timedelta(hours=2)).strftime('%m/%d/%Y %H:%M %p')
print("Rango de mensajes:", fecha_ini, "a", fecha_fin)

mensajes_filtrados = mensajes.Restrict(f"[ReceivedTime] >= '{fecha_ini}' AND [ReceivedTime] <= '{fecha_fin}'")
for msg in mensajes_filtrados:
    print('Asunto: ', msg.Subject)
    print('Fecha: ', msg.ReceivedTime)
    #print('Cuerpo: ', msg.Body)

    fecha_asunto = extraer_fecha(msg.Subject,
                                re.compile(rf"^WindCube Insights Fleet: New STA File from {id}\s+(\d{{4}}/\d{{2}}/\d{{2}})\s+(\d{{2}}[:\-]\d{{2}}[:\-]\d{{2}})$"),
                                 "windcube")
    
    print('Fecha asunto: ', fecha_asunto)




