# Control de correos meteorológicos

Este repositorio contiene un script en Python para **controlar la recepción de correos electrónicos de distintos sistemas meteorológicos** a través de Outlook y actualizar automáticamente un archivo de control con los resultados.  

El flujo principal es el siguiente:
1. Conexión a la cuenta de Outlook configurada.
2. Búsqueda de correos de remitentes conocidos (estaciones, windcube, zx, etc.).
3. Extracción de las fechas relevantes desde los asuntos de los correos.
4. Registro de la información en un archivo CSV y exportación a Excel con formato condicional.

---

## 📂 Estructura de archivos

- **`control_emails.py`** → Script principal.  
   - Conecta a Outlook, procesa los correos y actualiza los resultados en los archivos de salida.  
   - Puede ejecutarse indicando fecha de inicio y, opcionalmente, fecha final.  

- **`control_emails.csv`** → Archivo de entrada/salida.  
   - Contiene la lista de sistemas a controlar.  
   - Se actualiza con los resultados diarios (1 = recibido, 0 = no recibido).  

- **`control_emails.xlsx`** → Archivo de salida en Excel.  
   - Se genera a partir del CSV con formato condicional: verde para recibido y rojo para no recibido.  

- **`Instrucciones.txt`** → Documento con indicaciones específicas de uso y configuración.  

- **`.venv/`, `.vscode/`, `Desarrollo/`** → Carpetas auxiliares del entorno de desarrollo (no necesarias para la ejecución).  

- **`README.md`** → Este archivo.  

---

## 🛠️ Requisitos  

- Tener Python y los paquetes **`pandas`, `openpyxl`, `pywin32`** instalados.

- Tener Microsoft Outlook instalado y configurado en el sistema.

--- 

## 📒 Notas

- El script depende de que Outlook esté instalado y correctamente configurado con la cuenta indicada en el código.

- Los remitentes y patrones de asunto están predefinidos en el script.

- Es recomendable revisar el archivo Instrucciones.txt antes de la primera ejecución.

- Los archivos **`.venv/`**, **`.vscode/`** se usan únicamente en el entorno de desarrollo y no son necesarios para la ejecución.

- Archivos temporales de Windows como **`Thumbs.db`** no deben subirse al repositorio.

---

## ▶️ Uso

Desde la terminal, dentro de la carpeta del repositorio:

```bash
python control_emails.py YYYY-MM-DD [YYYY-MM-DD]
