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

# ---------------------------------------------------------------- #

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
conn_str = f'DRIVER={config('driver')};SERVER={config('server')};DATABASE={config('database')};UID={config('username')};PWD={config('password')}'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
cursor.execute("""
        select year(max(Fecha_de_operacion)) as [Year],
        month(max(Fecha_de_operacion)) as [Month],
        day(max(Fecha_de_operacion)) as [Day] 
        from volumen_restituido_petro
    """)
resultado = cursor.fetchone()
fecha_inicial = date(resultado.Year, resultado.Month, resultado.Day)
print(fecha_inicial)
fecha_final = date.today() - timedelta(days=1)
options = webdriver.ChromeOptions()
options.add_argument("--headless")
service = ChromeService(executable_path="Util/chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)
def verificar_elemento():
    if driver.find_elements(By.XPATH, "//b[text()='No se encontró información']"):
        pass
    else:
        datos_de_tabla_volumenrestituido = driver.find_element(By.XPATH,"//form[@action='/portal/sccli040/controlador']/following-sibling::table[1]")
        filas = datos_de_tabla_volumenrestituido.find_elements(By.TAG_NAME, 'tr')
        columnas = filas[0].find_elements(By.TAG_NAME, 'th')  
        data = []
        for fila in filas[1:]:
            celdas = fila.find_elements(By.TAG_NAME, 'td')
            fila_data = [celda.text for celda in celdas]
            data.append(fila_data)
        df = pd.DataFrame(data, columns=[celda.text for celda in columnas])
        df = df.iloc[:-2]
        nombres_columnas_df = df.columns.tolist()
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
            'Destino': 'Destino'}
        
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

        tabla_destino = 'Volumen_Restituido_Petro'
        for index, row in df.iterrows():
            query = f"""INSERT INTO {tabla_destino} 
                (Fecha_de_operacion, 
                Permiso_CRE, 
                Comprobante_de_carga, 
                Remision, 
                Producto, 
                Volumen_comprobante_de_carga, 
                Volumen_total_de_factura, 
                Volumen_a_restituir, 
                Cliente, 
                Destino) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
            cursor.execute(query, tuple(row)) 
        query_3 = ("""with C as (select  id, comprobante_de_carga, Remision, ROW_NUMBER() 
                   over (PARTITION BY comprobante_de_carga ORDER BY comprobante_de_carga desc) 
                   AS DUPLICADO FROM Volumen_Restituido) 
                   DELETE FROM C WHERE DUPLICADO >1""")
        cursor.execute(query_3)
        conn.commit()
while True:
    pagina_de_inicio = 'https://www.comercialrefinacion.pemex.com/portal/'
    pagina_volumen_restituido = 'https://www.comercialrefinacion.pemex.com/portal/sccli040/controlador?Destino=sccli040_01.jsp#'
    driver.get(pagina_de_inicio)
    driver.find_element(By.NAME,'usuario').send_keys(config('usuario_Pmx'))
    driver.find_element(By.NAME,'contrasena').send_keys(config('contrasena_pmx'))
    driver.find_element(By.NAME,'botonEntrar').send_keys(Keys.ENTER)      
    try:
        nombre_user_activo_pmx = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CLASS_NAME, 'textoEjecutivo')))
        break
    except:
        # print('Usuario ocupado, reintentando.')
        pass
for i in range((fecha_final - fecha_inicial).days + 1):
    fecha_a_validar = fecha_inicial + timedelta(days=i)
    mes = fecha_a_validar.strftime('%B').capitalize()
    dia = int(fecha_a_validar.day)
    driver.get(pagina_volumen_restituido)
    ventana_original = driver.current_window_handle
    # print(f"Trabajando con la fecha: {fecha_a_validar}, mes: {mes}, dia: {dia} .")
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
    verificar_elemento()
while True:
    try:
        driver.find_element(By.XPATH,"//ul[@id='nav']/li[4]/a[1]").click()
        validacion = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH,"//p[text()=' Contraseña : ']")))
        if validacion.text == 'Contraseña :':
            # print('Cierre de sesión completado con exito')
            break
        else:
            # print('No se ha podido cerrar la sesion')
            break
    except:
        pass

driver.quit()
conn.close()