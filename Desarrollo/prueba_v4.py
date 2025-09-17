''' 
Contexto: Cada día se envian correos a energias.renovables.es@dekra.com en los cuales se adjuntan los 
datos medidos por cada estación (LiDAR, SoDAR o TM). Cuando alguno de los sistemas falla no se envía el 
correo correspondiente. No hay forma de ver de forma directa si faltan correos y por tanto alguna de los 
sistemas está fallando. Se quiere automatizar un proceso por el cual se identifique si faltan correos o 
no y guardar en una hoja de cálculo un 1/0 por sistema y día además de un identificador del remitente.

El objetivo es tener un histórico de si se han recibido correos o no para cada sistema y día (filas=sistemas,
columnas=días)

Ejemplo, recibimos un correo de windcubeinsights@vaisala.info para la estación de Potrillo con fecha 2025-08-11
debemos guardar en la hoja de cálculo "1(windcube)" en la fila de Potrillo y en la columna del día 2025-08-11.
'''


'''
Asuntos de ejemplo y formato de fechas para cada remitente
estaciones.meteo -> LIDAR Punago-9_2025-08-12_00-10-00 ; YYYY-MM-DD_HH-MM-SS
windcube -> WindCube Insights Fleet: New STA File from WLS71497  2025/07/31  00:10:00 ; YYYY/MM/DD  HH-MM-SS
emailrelay -> LIDAR Villalube-6A_2025-08-11_00-10-00 ; YYYY-MM-DD_HH-MM-SS
zx -> Daily Data: Wind10_1148@Y2025_M08_D02.CSV (Averaged data) ; YYYYY_MMMM_DDD
estaciones.meteo (Olmillos) -> Ammonit Data Logger Meteo-40M D243094 Olmillos_1  (signed)
'''

########################################## 26/08/2025 ##########################################

import pandas as pd 
import win32com.client
import re 
import os

# Lectura de archivo .csv con datos de control de correos
df = pd.read_csv('control_emails.csv')
sistemas = df.iloc[:, 0].tolist()

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
nombre_cuenta = "energias.renovables.es@dekra.com"
store = outlook.Stores[nombre_cuenta]
# Acceso a la bandeja de entrada
bandeja_entrada = store.GetDefaultFolder(6)
# Acceso a la subcarpeta "Dades Meteo"
carpeta_dades_meteo = bandeja_entrada.Folders["Dades Meteo"]
# Items de la carpeta
#mensajes = bandeja_entrada.Items
mensajes = carpeta_dades_meteo.Items
mensajes.Sort("[ReceivedTime]", True)

# for i, msg in enumerate(mensajes): # Se muestran correctamente los mensajes
#     if i > 20:
#         break
#     print(msg.ReceivedTime, type(msg.ReceivedTime))

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


# -------- Procesar un único día --------

# Aquí definimos la fecha del día que queremos comprobar
fecha_objetivo = pd.Timestamp("2025-08-11")  # ejemplo, puedes cambiar la fecha
#fecha_objetivo = pd.Timestamp.now().normalize() # Formato 

fecha_str = fecha_objetivo.strftime("%Y-%m-%d")

# -------- Filtrar mensajes ±2 días usando Restrict --------
fecha_actual = fecha_objetivo.normalize()  # usar fecha objetivo como referencia
fecha_ini = (fecha_actual - pd.Timedelta(days=2)).strftime("%m/%d/%Y %H:%M:%S")
fecha_fin = (fecha_actual + pd.Timedelta(days=2)).strftime("%m/%d/%Y %H:%M:%S")

print("Rango de mensajes:", fecha_ini, "a", fecha_fin)

mensajes_filtrados = mensajes.Restrict(f"[ReceivedTime] >= '{fecha_ini}' AND [ReceivedTime] <= '{fecha_fin}'")

print(f"Número de mensajes en rango: {mensajes_filtrados.Count}")

# ------------------------- Procesar cada sistema -------------------------
resultados = []

for sistema in sistemas:

    id = df[df.iloc[:, 0] == sistema].iloc[0, 1]
    #print(id)
    remitente = df[df.iloc[:, 0] == sistema].iloc[0, 2]
    rem = ''
    #print(f"Procesando sistema: {sistema}, ID: {id}, Remitente: {remitente}")

    # Definir patrones según el remitente
    if remitente == 'estaciones.meteo@dekra-industrial.es':
        patron_asunto = re.compile(rf"^{id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
        rem = 'estacionesmeteo'
    elif remitente == 'windcubeinsights@vaisala.info':
        patron_asunto = re.compile(rf"^WindCube Insights Fleet: New STA File from {id}\s+(\d{{4}}/\d{{2}}/\d{{2}})\s+(\d{{2}}[:\-]\d{{2}}[:\-]\d{{2}})$")
        rem = 'windcube'
    elif remitente == 'emailrelay@konectgds.com':
        patron_asunto = re.compile(rf"^{id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
        rem = 'emailrelay'
    elif remitente == 'status@support.zxlidars.com':
        patron_asunto = re.compile(rf"^Daily Data: Wind10_{id}@Y(\d{{4}})_M(\d{{2}})_D(\d{{2}})\.CSV \(Averaged data\)$")
        rem = 'zx'
    elif remitente == 'estaciones.meteo@dekra-industrial.es' and id == 'Olmillos_1':
        patron_asunto = 'Ammonit Data Logger Meteo-40M D243094 Olmillos_1  (signed)'
        rem = 'estacionesmeteo (olmillos)'

    #print('rem: ', rem)

    valor = 0  # Valor por defecto: no se recibió correo de ese remitente

    for msg in mensajes_filtrados:

        # print('msg.Subject: ', msg.Subject)
        # print('msg.ReceivedTime: ', msg.ReceivedTime)
        # print('msg.Sender: ', msg.Sender)

        try:
            sender = msg.Sender.GetExchangeUser().PrimarySmtpAddress
        except:
            sender = msg.SenderEmailAddress

        if sender.lower() == remitente.lower():
            #print('sender=remitente= ', sender)
            if rem == 'estacionesmeteo (olmillos)':
                fecha_asunto_tmp = pd.Timestamp(msg.ReceivedTime).date()
                #print('Fecha Olmillos: ', fecha_asunto_tmp)
            else:
                fecha_asunto_tmp = extraer_fecha(msg.Subject, patron_asunto, rem)
            
            if not fecha_asunto_tmp:
                continue

            fecha_asunto = pd.to_datetime(fecha_asunto_tmp).date()

            if fecha_asunto == fecha_objetivo.date():
                valor = 1
                break
    
    # Guardamos los resultados para cada sistema
    resultados.append({
        "Sistema": sistema,
        "Remitente": remitente, 
        "Fecha": fecha_str,
        "Valor": valor
    })

# -------- Pasar resultados a DataFrame --------
df_result = pd.DataFrame(resultados)
print(df_result)

# -------- Actualizar control_emails.csv --------
output_file = "control_emails.csv"
fecha_col = fecha_str

# Si existe, cargar; si no, crear tabla base
if os.path.exists(output_file):
    tabla = pd.read_csv(output_file, index_col="Sistema")
else:
    tabla = pd.DataFrame(index=sistemas)
    tabla['Remitente'] = df.set_index("Sistema")["Remitente"]

# Añadir nueva columna si no existe
if fecha_col not in tabla.columns:
    tabla[fecha_col] = 0

# Actualizar valores del día
for r in resultados:
    tabla.loc[r["Sistema"], "Remitente"] = r["Remitente"]
    tabla.loc[r["Sistema"], fecha_col] = r["Valor"]

# Guardar actualizado
tabla.to_csv(output_file)
print("Archivo control_emails.csv actualizado")


