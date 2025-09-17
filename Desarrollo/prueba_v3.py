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

''' Regular expression operations (regex) -> Reconocer patrones en strings
- re.match(patron, texto) : Comprueba si el texto empieza con el patrón
- re.search(patron, texto) : Busca el patrón en cualquier parte del texto
- re.findall(patron, texto) : Devuelve todas las ocurrencias del patrón en el texto y las pone en una lista
- re.sub(patron, reemplazo, texto) : Reemplaza las ocurrencias del patrón en el texto
- re.compile(patron) : Compila el patrón para reutilizarlo varias veces de forma más eficiente
'''

'''
Asuntos de ejemplo y formato de fechas para cada remitente
estaciones.meteo -> LIDAR Punago-9_2025-08-12_00-10-00 ; YYYY-MM-DD_HH-MM-SS
windcube -> WindCube Insights Fleet: New STA File from WLS71497  2025/07/31  00:10:00 ; YYYY/MM/DD  HH-MM-SS
emailrelay -> LIDAR Villalube-6A_2025-08-11_00-10-00 ; YYYY-MM-DD_HH-MM-SS
zx -> Daily Data: Wind10_1148@Y2025_M08_D02.CSV (Averaged data) ; YYYYY_MMMM_DDD
estaciones.meteo (Olmillos) -> Ammonit Data Logger Meteo-40M D243094 Olmillos_1  (signed)
'''

########################################## 25/08/2025 ##########################################################################

import pandas as pd 
import win32com.client
import re 


df = pd.read_excel('sistemas_id_asuntos.xlsx')
sistemas = df.iloc[:, 0].tolist()
#print("Sistemas: ", sistemas)
#print(df)

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
nombre_cuenta = "energias.renovables.es@dekra.com"
store = outlook.Stores[nombre_cuenta]
bandeja_entrada = store.GetDefaultFolder(6)
mensajes = bandeja_entrada.Items
mensajes.Sort("[ReceivedTime]", True)


# Fecha actual
#fecha_actual = pd.Timestamp("2025-08-22").normalize() # Para pruebas
fecha_actual = pd.Timestamp.now().normalize() # Formato 
fecha_actual_ini = fecha_actual.date()  # Fecha actual a las 00:00:00
fecha_actual_fin = (fecha_actual + pd.Timedelta(days=1)).date()  # Un día después

print("Fecha inicial: ", fecha_actual_ini)
print("Fecha final: ", fecha_actual_fin)


# Definir rango de fechas en formato Outlook
inicio = fecha_actual.strftime("%m/%d/%Y %H:%M:%S")
fin = (fecha_actual + pd.Timedelta(days=1)).strftime("%m/%d/%Y %H:%M:%S")

# Filtrar mensajes entre esas fechas
filtro = f"[ReceivedTime] >= '{inicio}' AND [ReceivedTime] < '{fin}'"
mensajes_filtrados = mensajes.Restrict(filtro)

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


resultados = []

for sistema in sistemas:

    print(sistema)

    id = df[df.iloc[:, 0] == sistema].iloc[0, 1]
    remitente = df[df.iloc[:, 0] == sistema].iloc[0, 2]
    rem = ''

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

    #print(patron_asunto)
    
    # Valor por defecto: no se recibió correo de ese remitente
    fecha_asunto = str(fecha_actual_ini)
    out = f"0 ({rem})"


    for msg in mensajes_filtrados:
        #print(msg.Subject)
        #print(pd.Timestamp(msg.ReceivedTime).date())
        #print(msg.SenderEmailAddress)

        try:
            sender = msg.Sender.GetExchangeUser().PrimarySmtpAddress
        except:
            sender = msg.SenderEmailAddress

        #print(sender)

        if sender.lower() == remitente.lower():
            if rem == 'estacionesmeteo (olmillos)':
                fecha_asunto_tmp = pd.Timestamp(msg.ReceivedTime).date()
            else:
                fecha_asunto_tmp = extraer_fecha(msg.Subject, patron_asunto, rem)
                #if contiene_elementos(msg.Subject) == "sí":
                if fecha_asunto_tmp: # No es None ni una cadena vacía
                    valor = f"1 ({rem})"
                break 
        
    # Crear data frame con columnas = fecha_asunto, filas = sistema y valores = out
    resultados.append({
        "Sistema": sistema,
        "Fecha": fecha_asunto if fecha_asunto else str(fecha_actual.date()),
        "Valor": out
    })

df_resultados = pd.DataFrame(resultados)
tabla = df_resultados.pivot(index="Sistema", columns="Fecha", values="Valor")

print(tabla)

tabla.to_csv("resultados_correos.csv", index=True)