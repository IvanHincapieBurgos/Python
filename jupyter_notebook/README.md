# Google Colab

El cómo llegue a conocerlo es debido a un proyecto, que a su vez me llevo a una búsqueda por la optimización de las tareas manuales. Siendo una gran ganacia de tiempo para el equipo de Customer Experience (CX). A través del aprendizaje y aplicación de Google Colab, logré transformar un proceso repetitivo en una solución automatizada y eficiente.

## Contexto del Problema
Las tareas consistían en gestionar los cupones habilitados para el equipo de CX, que involucraban los siguientes pasos:

1. Acceder a Google Drive, donde se encontraban dos carpetas con archivos tipo Google Sheets (GSheet).
2. En cada archivo, revisar cada pestaña, que representaba un valor de cupón (por ejemplo, $5, $10, etc.).
3. Copiar manualmente la única columna de cupones de cada pestaña.
4. Crear una nueva columna indicando el valor del cupón basado en el nombre de la pestaña.
5. Repetir este proceso para todas las pestañas y consolidar los datos en un solo archivo.
6. Comparar esta información con otra Gshhet que contenía cupones de otros departamentos para evitar duplicados.

Este proceso manual no solo era tedioso, sino que también era propenso a errores humanos.

## Solución: Automatización con Google Colab

Motivado por la necesidad de optimizar esta tarea, aprendí a utilizar **Google Colab**. Con la automatización:

- Se accede directamente a las carpetas de Google Drive.
- Se procesan todas las pestañas de los archivos de manera automática.
- Se genera una nueva columna con el valor del cupón basado en el nombre de cada pestaña.
- Se consolidan todos los datos en un solo archivo.
- Se valida la información comparándola con los cupones de otros departamentos para evitar duplicados.

## Resultados

Con esta solución:
- Eliminé el riesgo de errores asociados con la gestión manual de datos.
- Aumenté la eficiencia y productividad en la preparación de cupones para los asesores de CX.
- Reduje significativamente el tiempo necesario para completar esta tarea.

---
¡Espero que esta iniciativa también inspire a automatizar tareas repetitivas y a compartir tus aprendizajes con la comunidad! 🚀
---
## Cosas Útiles 

### ¿Cómo puedo hacer para que Google Colab se conecte con mi Drive? (Conexión a Google Drive)
1. Ejecuta el siguiente código para montar Google Drive en tu entorno de Colab:
   ```python
   from google.colab import drive
   drive.mount('/content/drive')
   ```
2. Una vez montado, puedes acceder a los archivos de Drive como si fueran parte del sistema de archivos local.

### ¿Cómo puedo leer un Archivo GSheet y Convertirlo en un DataFrame?
1. Autenticar y otorgar permisos para conectar con Google Sheets:
   ```python
   # Authenticate and grant permissions to connect with Google Sheets
   from google.colab import auth
   auth.authenticate_user()

   # Import necessary libraries for Google Sheets integration
   from gspread_dataframe import set_with_dataframe
   import gspread
   from google.auth import default
   creds, _ = default()
   gc = gspread.authorize(creds)

   # Libraries for data manipulation and analysis
   import pandas as pd
   ```

2. Conectar al archivo de Google Sheets y convertirlo en un DataFrame:
   ```python
   cupons = 'ID_del_Archivo' # Ve a la URL del archivo y copia lo que se encuentre después del "https://docs.google.com/spreadsheets/d/" y antes del slash que cierra esa parte.

   gsheet = gc.open_by_key(cupons).worksheet('Sheet')
   
   # Convertir la hoja de cálculo en un DataFrame
   df = pd.DataFrame(gsheet.get_all_values()[1:], columns=gsheet.get_all_values()[0])
   print(df.head())
   ```
