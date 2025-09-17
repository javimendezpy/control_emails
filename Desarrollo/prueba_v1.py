import win32com.client
import datetime
import pandas as pd
import re
import os

# === CONFIGURACIÓN ===
fecha_objetivo = datetime.date.today()  # Cambia para probar otro día
nombre_cuenta = "energias.renovables.es@dekra.com"
excel_sistemas = r"Z:\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\CorreosDatosEstaciones\sistemas_id_asuntos.xlsx"
csv_historico = r"Z:\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\CorreosDatosEstaciones\historico_correos.csv"

# === LEER SISTEMAS ===
df_sistemas = pd.read_excel(excel_sistemas)
columnas_esperadas = {"Sistema", "Identificador", "Emisor", "Formato"}
if not columnas_esperadas.issubset(df_sistemas.columns):
    raise ValueError(f"El Excel debe tener columnas: {columnas_esperadas}")

emisores = df_sistemas["Emisor"].unique()

# === LEER O CREAR HISTÓRICO ===
if os.path.exists(csv_historico):
    df_hist = pd.read_csv(csv_historico)
else:
    df_hist = pd.DataFrame({"Sistema": df_sistemas["Sistema"].unique()})

# === CONECTAR A OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print("\n📬 Cuentas de Outlook detectadas:")
for store in outlook.Stores:
    print(" -", store.DisplayName)

try:
    store = outlook.Stores[nombre_cuenta]
except Exception:
    raise ValueError(f"No se encontró la cuenta '{nombre_cuenta}'. Verifica el nombre en la lista anterior.")

bandeja_entrada = store.GetDefaultFolder(6)
mensajes = bandeja_entrada.Items
mensajes.Sort("[ReceivedTime]", True)

# === CREAR DICCIONARIO DE PATRONES POR EMISOR ===
patrones_por_emisor = {}

for emisor in emisores:
    filas_emisor = df_sistemas[df_sistemas["Emisor"].str.lower().str.strip() == emisor.lower().strip()]
    if filas_emisor.empty:
        patrones_por_emisor[emisor] = None
        print(f"⚠ No se encontró el emisor en el Excel: {emisor}")
        continue

    # Buscar primera fila con patrón válido
    filas_validas = filas_emisor["Formato"].dropna() # Elimina NaN
    filas_validas = filas_validas[filas_validas.astype(str).str.strip() != ""]
    if filas_validas.empty:
        patrones_por_emisor[emisor] = None
        print(f"⚠ Emisor '{emisor}' no tiene patrón en la columna 'Formato'")
        continue

    patron_texto = filas_validas.iloc[0]
    try:
        patrones_por_emisor[emisor] = re.compile(patron_texto)
    except re.error as e:
        patrones_por_emisor[emisor] = None
        print(f"❌ Error en regex para emisor '{emisor}': {e}")

# === PREPARAR COLUMNA PARA ESTA FECHA Y EMISOR ===
nueva_col_names = {emisor: f"{fecha_objetivo}_{emisor}" for emisor in emisores}
for col in nueva_col_names.values():
    if col not in df_hist.columns:
        df_hist[col] = 0  # Si la columna no existe, la crea con valores 0

# === BÚSQUEDA EN BANDEJA DE ENTRADA ===
contador = 0
for mensaje in mensajes:
    try:
        asunto = mensaje.Subject
        remitente = mensaje.SenderEmailAddress.lower().strip()

        if remitente not in emisores:
            continue

        patron_asunto = patrones_por_emisor.get(remitente)
        if patron_asunto is None:
            continue

        match = patron_asunto.match(asunto)
        if not match:
            continue

        # Si hay fecha en el asunto
        if len(match.groups()) >= 2:
            try:
                fecha_asunto = datetime.datetime.strptime(match.group(2), "%Y-%m-%d").date()
                if fecha_asunto != fecha_objetivo:
                    continue
            except ValueError:
                print(f"⚠ Fecha inválida en asunto: '{asunto}' de {remitente}")
                continue

        sistema_id = match.group(1).strip()

        fila_sistema = df_sistemas[
            (df_sistemas["Identificador"] == sistema_id) &
            (df_sistemas["Emisor"].str.lower().str.strip() == remitente)
        ]
        if not fila_sistema.empty:
            sistema_nombre = fila_sistema["Sistema"].iloc[0]
            df_hist.loc[df_hist["Sistema"] == sistema_nombre, nueva_col_names[remitente]] = 1

        contador += 1
        if contador >= 200:
            break

    except AttributeError:
        continue

# === GUARDAR HISTÓRICO ===
df_hist.to_csv(csv_historico, index=False)
print(f"\n✅ Histórico actualizado en: {csv_historico}")

