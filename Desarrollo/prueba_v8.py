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

########################################## 12/09/2025 ##########################################

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
bandeja_entrada = store.GetDefaultFolder(6) # Acceso a la bandeja de entrada
carpeta_dades_meteo = bandeja_entrada.Folders["Dades Meteo"] # Acceso a la subcarpeta "Dades Meteo"
mensajes = carpeta_dades_meteo.Items # Items de la carpeta
# mensajes = bandeja_entrada.Items # Items de la carpeta
mensajes.Sort("[ReceivedTime]", True)

################################################################################################

def extraer_remitente(remitente: str, id: str) -> str:

    ' Esta función extrae un identificador según el remitente y su id '
    if remitente == 'estaciones.meteo@dekra-industrial.es' and id == 'Olmillos_1':
        return 'estacionesmeteo (olmillos)'
    elif remitente == 'estaciones.meteo@dekra-industrial.es':
        return 'estacionesmeteo'
    elif remitente == 'windcubeinsights@vaisala.info':
        return 'windcube'
    elif remitente == 'emailrelay@konectgds.com':
        return 'emailrelay'
    elif remitente == 'status@support.zxlidars.com':
        return 'zx'
    return ''


def extraer_patron(remitente, id):

    ' Esta función extrae el patrón regex según el remitente y su id '

    ''' -------- Asuntos de ejemplo y formato de fechas para cada remitente --------
    estaciones.meteo -> LIDAR Punago-9_2025-08-12_00-10-00 ; YYYY-MM-DD_HH-MM-SS
    windcube -> WindCube Insights Fleet: New STA File from WLS71497  2025/07/31  00:10:00 ; YYYY/MM/DD  HH-MM-SS
    emailrelay -> LIDAR Villalube-6A_2025-08-11_00-10-00 ; YYYY-MM-DD_HH-MM-SS
    zx -> Daily Data: Wind10_1148@Y2025_M08_D02.CSV (Averaged data) ; YYYYY_MMM_DDD
    estaciones.meteo (Olmillos) -> Ammonit Data Logger Meteo-40M D243094 Olmillos_1  (signed) '''

    if remitente == 'estaciones.meteo@dekra-industrial.es' and id == 'Olmillos_1':
        return None # Olmillos no utiliza fecha en el asunto
    elif remitente == 'estaciones.meteo@dekra-industrial.es':
        return re.compile(rf"^{id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
    elif remitente == 'windcubeinsights@vaisala.info':
        return re.compile(rf"^WindCube Insights Fleet: New STA File from {id}\s+(\d{{4}}/\d{{2}}/\d{{2}})\s+(\d{{2}}[:\-]\d{{2}}[:\-]\d{{2}})$")
    elif remitente == 'emailrelay@konectgds.com':
        return re.compile(rf"^{id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
    elif remitente == 'status@support.zxlidars.com':
        return re.compile(rf"^Daily Data: Wind10_{id}@Y(\d{{4}})_M(\d{{2}})_D(\d{{2}})\.(?:CSV|ZPH) \(Averaged data\)$")
    return None


def extraer_fecha(asunto: str, patron, remitente: str, received_time=None) -> str | None:

    ' Esta función extrae la fecha del asunto del correo en función del identificador del remitente '

    ''' Devuelve la fecha de datos (YYYY-MM-DD) según remitente.
    Para estaciones.meteo/emailrelay → fecha del asunto -1 día
    Para windcube/zx → fecha del asunto directamente
    Para Olmillos → fecha de recepción -1 día
    '''

    if remitente == 'estacionesmeteo (olmillos)':
        if received_time is None:
            return None
        # Forzar datetime naive (sin tz)
        if hasattr(received_time, "replace"):
            received_time = received_time.replace(tzinfo=None)
        fecha = (pd.Timestamp(received_time) - pd.Timedelta(days=1)).date()
        return str(fecha)
    
    if not patron:
        return None
    
    m = patron.search(asunto)
    if not m:
        return None
    
    if remitente in ["estacionesmeteo", "emailrelay"]:
        fecha = pd.to_datetime(m.group(1)).date()-pd.Timedelta(days=1)
        return str(fecha)  # ya viene YYYY-MM-DD
    
    elif remitente == "windcube":
        fecha = pd.to_datetime(m.group(1).replace("/", "-")).date()
        return str(fecha)  # YYYY/MM/DD -> YYYY-MM-DD
    
    elif remitente == "zx":
        fecha = pd.to_datetime(f"{m.group(1)}-{m.group(2)}-{m.group(3)}").date()
        return str(fecha)
    return None


################################################################################################

# Aquí definimos la fecha del día que queremos comprobar

# Pasar una fecha manual
fecha_referencia = pd.to_datetime("2025-09-04").date()

print(f"Fecha de referencia (día de datos): {fecha_referencia}")

dia_siguiente = fecha_referencia + pd.Timedelta(days=1)

print(f'Esperamos encontrar los archivos de datos del día {fecha_referencia} en el día {dia_siguiente}.')
fecha_ini = pd.Timestamp(fecha_referencia).strftime('%d/%m/%Y 00:00 AM')
fecha_fin = pd.Timestamp(dia_siguiente).strftime('%d/%m/%Y 23:59 PM')

print("Rango de tiempo de recepción de mensajes:", fecha_ini, "a", fecha_fin)

# -------- Filtrar mensajes ±2 días usando Restrict --------
mensajes_filtrados = mensajes.Restrict(f"[ReceivedTime] >= '{fecha_ini}' AND [ReceivedTime] <= '{fecha_fin}'")

print(f"Número de mensajes en rango: {mensajes_filtrados.Count}")

################################################################################################

# Bucle sobre todos los sistemas 
resultados = []

for sistema in sistemas:

    remitente = df[df.iloc[:, 0] == sistema].iloc[0, 1]
    id = df[df.iloc[:, 0] == sistema].iloc[0, 2]

    #print(f"\n\n\n Procesando SISTEMA: {sistema}, con ID: {id} y REMITENTE: {remitente} \n\n")

    rem = extraer_remitente(remitente, id)
    patron_asunto = extraer_patron(remitente, id)
    #print(patron_asunto)

    valor = 0  # Valor por defecto: no se recibió correo de ese remitente

    for msg in mensajes_filtrados:

        # print('msg.Subject: ', msg.Subject)
        # print('msg.ReceivedTime: ', msg.ReceivedTime)
        # print('msg.Sender: ', msg.Sender)

        try:
            sender = msg.Sender.GetExchangeUser().PrimarySmtpAddress
        except:
            sender = msg.SenderEmailAddress

        # sender = msg.SenderEmailAddress

        # print('Remitente: ', sender)

        if sender.lower() == remitente.lower():
            # print('Se ha encontrado un correo con remitente igual al del sistema: ', sender)
            fecha_asunto_tmp = extraer_fecha(
                msg.Subject, 
                patron_asunto, 
                rem, 
                received_time=msg.ReceivedTime
            )
            #print('Fecha asunto: ', fecha_asunto_tmp)
            if not fecha_asunto_tmp:
                continue

            fecha_asunto = pd.to_datetime(fecha_asunto_tmp).tz_localize(None).date()
            #print('Fecha asunto (formateada): ', fecha_asunto)

            if fecha_asunto == fecha_referencia:
                valor = 1
                break
    
    # Guardamos los resultados para cada sistema
    resultados.append({
        "Sistema": sistema,
        "Remitente": remitente, 
        "Fecha": fecha_referencia,
        "Valor": valor
    })

# -------- Pasar resultados a DataFrame --------
df_result = pd.DataFrame(resultados)
print(df_result)

# -------- Actualizar control_emails.csv --------
output_file = "control_emails.csv"
fecha_col = str(fecha_referencia)

# Si existe, cargar; si no, crear tabla base
if os.path.exists(output_file):
    tabla = pd.read_csv(output_file, index_col="Sistema")
else:
    tabla = pd.DataFrame(index=sistemas)
    tabla['Remitente'] = df.set_index("Sistema")["Remitente"]

# Si la fecha no existe escribirla y rellenar con 0
if fecha_col not in tabla.columns:
    tabla[fecha_col] = 0
    
    # Actualizar valores del día
    for r in resultados:
        # tabla.loc[r["Sistema"], "Remitente"] = r["Remitente"]
        tabla.loc[r["Sistema"], fecha_col] = r["Valor"]

    print(f"Se ha añadido la columna {fecha_col} y actualizado con los resultados.")

# Actualizar valores del día
for r in resultados:
    tabla.loc[r["Sistema"], "Remitente"] = r["Remitente"]
    tabla.loc[r["Sistema"], fecha_col] = r["Valor"]

# Reordenar columnas: dejar fijas y fechas descendentes
cols_fijas = [c for c in tabla.columns if not re.match(r"^\d{4}-\d{2}-\d{2}$", c)]
cols_fechas = sorted(
    [c for c in tabla.columns if re.match(r"^\d{4}-\d{2}-\d{2}$", c)],
    reverse=True
)
tabla = tabla[cols_fijas + cols_fechas]

 #print(tabla)

# Guardar control_emails.csv actualizado
tabla.to_csv(output_file)
tabla.to_excel('control_emails.xlsx', index=False)
# print("Archivo control_emails (.xlsx/.csv) actualizado")


