# Importación de módulos de Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
# Importación de otros módulos
import pyodbc
import pandas as pd
from decouple import config
from datetime import date, timedelta
import locale

#  ------ Configuración de acceso y conexiona SQL Server ------ 
conn_str = f'DRIVER={config('driver')};SERVER={config('server')};DATABASE={config('database')};UID={config('username')};PWD={config('password')}'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ------ Configuracion selenium ------ 
options = webdriver.ChromeOptions()
# options.add_argument("--headless")
service = ChromeService(executable_path="C:/Program Files/Google/Chrome/Application/chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)

while True:
    # ------ Inicio de la navegacion ------ 
    pagina_de_inicio = 'https://www.comercialrefinacion.pemex.com/portal/'
    pagina_volumen_restituido = 'https://www.comercialrefinacion.pemex.com/portal/sccli040/controlador?Destino=sccli040_01.jsp#'
    driver.get(pagina_de_inicio)
    driver.find_element(By.NAME,'usuario').send_keys(config('usuario_Pmx'))
    driver.find_element(By.NAME,'contrasena').send_keys(config('contrasena_pmx'))
    driver.find_element(By.NAME,'botonEntrar').send_keys(Keys.ENTER)
            
#------ En este momento ya hicimos login en la pagina ------ 
    try:
        nombre_user_activo_pmx = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CLASS_NAME, 'textoEjecutivo')))
        break
    except: 
        pass

#------ Va a el reporte de volumen restituido ------ 

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
fecha_actual_menos_uno = date.today() - timedelta(days=1)
mes = fecha_actual_menos_uno.strftime('%B').capitalize()
dia = int(fecha_actual_menos_uno.day)
driver.get(pagina_volumen_restituido)
ventana_original = driver.current_window_handle
# Agregar los elementos para cambiar de pagina en el calendario 
driver.find_element(By.XPATH,"//img[@src='/portal/imagenes/calendario.gif']").click()
driver.switch_to.window(driver.window_handles[1])
select_mes = Select(driver.find_element(By.NAME, "selMonth"))
select_mes.select_by_visible_text(mes)
driver.find_element(By.LINK_TEXT,str(dia)).click()
driver.switch_to.window(ventana_original)
driver.find_element(By.XPATH,"(//img[@src='/portal/imagenes/calendario.gif'])[2]").click()
driver.switch_to.window(driver.window_handles[1])
select_mes = Select(driver.find_element(By.NAME, "selMonth"))
select_mes.select_by_visible_text(mes)
driver.find_element(By.LINK_TEXT,str(dia)).click()
driver.switch_to.window(ventana_original)
driver.find_element(By.XPATH,"//input[@value='Consultar']").click()
datos_de_tabla_volumenrestituido = driver.find_element(By.XPATH,"//form[@action='/portal/sccli040/controlador']/following-sibling::table[1]")

# Obtener datos de la tabla y dividirlos en filas y columnas
filas = datos_de_tabla_volumenrestituido.find_elements(By.TAG_NAME, 'tr')
# Se asume que la primera fila contiene encabezados
columnas = filas[0].find_elements(By.TAG_NAME, 'th')  

# Crear un DataFrame de pandas con los datos de la tabla
data = []
for fila in filas[1:]:
    celdas = fila.find_elements(By.TAG_NAME, 'td')
    fila_data = [celda.text for celda in celdas]
    data.append(fila_data)


df = pd.DataFrame(data, columns=[celda.text for celda in columnas])
# Se descartan las dos ultimas filas ya que no contienen informacion
df = df.iloc[:-2]

while True:
    try:
        # A partir de esta parte se procede con el cierre de sesion
        driver.find_element(By.XPATH,"//ul[@id='nav']/li[4]/a[1]").click()
        validacion = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH,"//p[text()=' Contraseña : ']")))

        if validacion.text == 'Contraseña :':
            print('Cierre de sesión completado con exito')
            break
        else:
            print('No se ha podido cerrar la sesion')
        break
    except:
        pass

driver.quit()

# Se crea una lista con los nombres de los encabezados de las columnas.
nombres_columnas_df = df.columns.tolist()

# Se crea un diccionario con los nombres de las columnas del DataFrame
#   y el nombre que deben de llevar en sql
mapeo_columnas = {
    'Fecha de operación': 'Fecha_de_operacion',
    'Permiso CRE': 'Permiso_CRE',
    'Comprobante de carga': 'Comprobante_de_carga',
    'Remisión': 'Remision',
    'Producto': 'Producto',
    'Volumen comprobante de carga(lts)': 'Volumen_comprobante_de_carga',
    'Volumen total de factura(lts)': 'Volumen_total_de_factura',
    'Volumen a restituir (lts)': 'Volumen_a_restituir',
    'Cliente': 'Cliente',
    'Destino': 'Destino'
}

# Se renombrar columnas en el DataFrame y se establece el tipo de dato para cada columna
df.rename(columns=mapeo_columnas, inplace=True)
df['Fecha_de_operacion'] = pd.to_datetime(df['Fecha_de_operacion'], dayfirst= True)
df['Fecha_de_operacion'] = df['Fecha_de_operacion'].dt.strftime("%Y-%m-%d")
df['Permiso_CRE'] = df['Permiso_CRE'].astype(str)
df['Comprobante_de_carga'] = df['Comprobante_de_carga'].astype(str)
df['Remision'] = df['Remision'].astype(str)
df['Producto'] = df['Producto'].astype(str)
df['Volumen_comprobante_de_carga'] = df['Volumen_comprobante_de_carga'].apply(lambda x: float(x.replace(',', '')) if x.replace(',', '').replace('.', '').isdigit() else None)
df['Volumen_total_de_factura'] = df['Volumen_total_de_factura'].apply(lambda x: float(x.replace(',', '')) if x.replace(',', '').replace('.', '').isdigit() else None)
df['Volumen_a_restituir'] = df['Volumen_a_restituir'].apply(lambda x: float(x.replace(',', '')) if x.replace(',', '').replace('.', '').isdigit() else None)
df['Cliente'] = df['Cliente'].astype(str)
df['Destino'] = df['Destino'].astype(str)

# Ahora, los nombres de las columnas en el DataFrame coinciden con los de la tabla SQL Server
# Agregar datos a la tabla Volumen_Restituido en SQL Server
tabla_destino = 'Volumen_Restituido_Petro'
for index, row in df.iterrows():
    # Construir la consulta de inserción
    query = f"INSERT INTO {tabla_destino} (Fecha_de_operacion, Permiso_CRE, Comprobante_de_carga, Remision, Producto, Volumen_comprobante_de_carga, Volumen_total_de_factura, Volumen_a_restituir, Cliente, Destino) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    # Ejecutar la consulta con los valores de la fila actual
    cursor.execute(query, tuple(row))

# ------ Elimina los registros duplicados (Si existen) ------ 
query_3 = ("with C as (select  id, comprobante_de_carga, Remision, ROW_NUMBER() over (PARTITION BY comprobante_de_carga ORDER BY comprobante_de_carga desc) AS DUPLICADO FROM Volumen_Restituido) DELETE FROM C WHERE DUPLICADO >1")
cursor.execute(query_3)

# Confirmar los cambios en la base de datos
conn.commit()

# Cerrar la conexión a SQL Server
conn.close()
print('Finalizado')