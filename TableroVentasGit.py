from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import os
import pandas as pd
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText


# Definir las rutas de los archivos y el nombre del tablero en Power BI
reporte_diario_path = r"C:\Users\Magellan Banyuls\Downloads\ReporteDiario.xlsx"
power_bi_tablero_url = "URL_del_tablero_en_Power_BI"
opciones = Options()
opciones.headless = False  # Para ejecutar el navegador en modo sin cabeza (headless)

def descargar_archivo(driver, url, usuario, contrasena, nombre_archivo):
    driver.get(url)
    time.sleep(3)
    
    # Ingresar usuario y contraseña
    input_usuario = driver.find_element(By.ID, "user_email")  # Reemplazar por el nombre correcto del campo de usuario
    input_contrasena = driver.find_element(By.ID,"user_password")  # Reemplazar por el nombre correcto del campo de contraseña
    input_usuario.send_keys(str(usuario))
    input_contrasena.send_keys(str(contrasena))
    time.sleep(1)
    boton_iniciar = driver.find_element(By.XPATH, '/html/body/div[2]/div/form/div[2]/input')
    boton_iniciar.click()
    
    # Esperar a que se cargue la página correctamente
    time.sleep(2)
    # Entrar a solicitudes
    boton_solicitudes= driver.find_element(By.XPATH,'//*[@id="main-menu"]/li[4]/a')
    boton_solicitudes.click()
    time.sleep(0.5)
    # Entrar a todas las solicitudes
    boton_solicitudes= driver.find_element(By.XPATH,'//*[@id="main-menu"]/li[4]/ul/li[1]/a')
    boton_solicitudes.click()
    time.sleep(0.5)
    # Click en filtros
    boton_filtros = driver.find_element(By.XPATH, '//*[@id="page-inner"]/div/div/div[1]/div[1]/h3')
    boton_filtros.click()
    time.sleep(0.5)

    # Click en fecha desde
    boton_fecha_desde= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[1]/div/div/input')
    boton_fecha_desde.click()
    boton_fecha_desde= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[1]/div/div/input')
    boton_fecha_desde.send_keys(Keys.DELETE)
    boton_fecha_desde= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[1]/div/div/input')
    boton_fecha_desde.send_keys(fecha_desde)
    # Click fecha hasta
    boton_fecha_hasta= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[2]/div/div/input')
    boton_fecha_hasta.click()
    boton_fecha_hasta= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[2]/div/div/input')
    boton_fecha_hasta.send_keys(Keys.DELETE)
    boton_fecha_hasta= driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/div[3]/div[2]/div/div/input')
    boton_fecha_hasta.send_keys(fecha_hasta)

    time.sleep(2)
    # Click en buscar
    boton_buscar = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div[2]/form/input[3]')
    boton_buscar.click()
    time.sleep(1)

    # Click en el boton de descarga
    boton_descargar = driver.find_element(By.XPATH, '//*[@id="page-inner"]/div/div/div[33]/div/a')
    boton_descargar.click()
    # Esperar a que se descargue
    time.sleep(10)

    # Transformarlo en csv

    ruta_descarga = r"C:\Users\Magellan Banyuls\Downloads\solicitudes.xlsx"  # Reemplazar por la ruta correcta de descarga
    df = pd.read_excel(ruta_descarga)  # Leer el archivo Excel
    df.to_csv(ruta_descarga.replace('.xlsx', '.csv'), index=False, sep=";")
    
  
    # Obtener la ruta del archivo descargado
    ruta_descarga_nueva = r"C:\Users\Magellan Banyuls\Downloads\solicitudes.csv"  # Reemplazar por la ruta correcta de descarga
    nuevo_nombre = f"C:\\Users\\Magellan Banyuls\\Downloads\\{nombre_archivo}.csv"

    # Borrar solicitud
    os.remove(ruta_descarga)
    # Verificar si ya existe un archivo con el mismo nombre y sobrescribirlo
    if os.path.exists(nuevo_nombre):
        os.remove(nuevo_nombre)
    os.rename(ruta_descarga_nueva, nuevo_nombre)





def main():
    # Iniciar el driver de Selenium
    driver_path = r"C:\Users\Magellan Banyuls\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"  # Ruta del controlador del navegador
    selenium_service = Service(driver_path)
    driver = webdriver.Chrome(service=selenium_service, options=opciones)

    
    # Usuario y contraseña
    usuario_customer = "Aqui va el usuario"
    contrasena_customer = "Aqui va la contraseña"
    usuario_2= "Aqui va el segundo usuario"
    contrasena_2="Aqui va la segunda contraseña"

    # Descargar el primer archivo y cambiarle el nombre
    descargar_archivo(driver, "https://customercorner.net/", usuario_customer, contrasena_customer, "Reporte1")
    
    # Descargar el segundo archivo y cambiarle el nombre
    descargar_archivo(driver, "https://ventas.osar.com.ar", usuario_2, contrasena_2, "Reporte2")

    # Cerrar el navegador
    driver.quit()

    if os.path.exists(reporte_diario_path):
        os.remove(reporte_diario_path)

    
    # Leer los csv
    df1 = pd.read_csv(r"C:\Users\Magellan Banyuls\Downloads\Reporte1.csv", sep=";", engine='python', na_values=[""])
    df2 = pd.read_csv(r"C:\Users\Magellan Banyuls\Downloads\Reporte2.csv", sep=";", engine='python', na_values=[""])

    # Concatenar los DataFrames
    df_concatenado = pd.concat([df1, df2])
    columnas_requeridas = [
        'Fecha de Venta',
        'Producto',
        'Precio',
        'Cantidad de portas',
        'Estado de Solicitud',
        'Provincia',
        'Ciudad',
        'Vendedor',
        'Fecha de Portación'
        ]

    # Seleccionar solo las columnas requeridas del DataFrame original
    df = df_concatenado[columnas_requeridas]

    # Establecer comision
    comision = float(2.2)

    # Pasar a formatos de fecha

    df['Fecha de Portación'] = pd.to_datetime(df['Fecha de Portación'], format='%d/%m/%Y', errors='coerce')
    df['Fecha de Venta'] = pd.to_datetime(df['Fecha de Venta'], format='%d/%m/%Y', errors='coerce')


    # Crear las columnas 'Dia' y 'Mes' a partir de la columna 'Fecha de Venta' y 'Fecha de portacion' respectivamente

    df['Dia'] = df['Fecha de Venta'].dt.day
    df['Mes'] = df['Fecha de Portación'].dt.month

    # Rellenar los meses con fecha de venta en el de portacion,esto es para tener una diferenciacion de ventas brutas y netas.

    df['Mes'].fillna(df['Fecha de Venta'].dt.month, inplace=True)


    # Filtrar el DataFrame para mantener solo los datos de los últimos 3 meses
    fecha_actual = pd.to_datetime('today')
    fecha_tres_meses_atras = fecha_actual - pd.DateOffset(months=3)
    df = df[pd.to_datetime(df['Fecha de Venta'], format='%d/%m/%Y') >= fecha_tres_meses_atras]


    # Eliminar el guion ("-") al final de los valores en la columna 'Precio'
    df['Precio'] = df['Precio'].str.replace(r'-$', '', regex=True)

    # Eliminar el prefijo "M " de los valores en la columna "Vendedor"
    df['Vendedor'] = df['Vendedor'].str.replace('^M ', '', regex=True)

    # Convertir la columna 'Precio' a tipo numérico
    df['Precio'] = pd.to_numeric(df['Precio'], errors='coerce')

    # Crear la columna 'Ventas Totales' que suma la cantidad de portas
    df['Ventas Totales'] = df['Cantidad de portas']

    # Crear la columna 'Ventas Brutas' que es la multiplicación de 'Cantidad de portas' por 'Precio'
    df['Ventas Brutas'] = df['Cantidad de portas'] * df['Precio'] * comision

    # Utilizar .loc para asignar los valores de Ventas Netas
    df.loc[df['Estado de Solicitud'] == 'APROBADA', 'Ventas Netas'] = df.loc[df['Estado de Solicitud'] == 'APROBADA', 'Precio'] * df.loc[df['Estado de Solicitud'] == 'APROBADA', 'Cantidad de portas'] * comision



    # Agrupar por mes y sumar las ventas brutas y netas para obtener el total por mes
    df_meses = df.groupby('Mes').agg({'Ventas Brutas': 'sum', 'Ventas Netas': 'sum'})

    # Calcular la suma total de la columna 'Ventas Brutas' y asignarla a la nueva columna 'Total Ventas Brutas'
    suma_ventas_brutas = df['Ventas Brutas'].sum() * comision
    df_meses['Total Ventas Brutas'] = suma_ventas_brutas

    # Calcular la suma total de la columna 'Ventas Netas' y asignarla a la nueva columna 'Total Ventas Netas'
    suma_ventas_netas = df['Ventas Netas'].sum() * comision
    df_meses['Total Ventas Netas'] = suma_ventas_netas


    # Unir el DataFrame df con el DataFrame df_meses utilizando merge
    df = df.merge(df_meses, on='Mes', how='left')

    # Agregar la columna 'Objetivo' con el valor 35 en la primera fila
    df['Objetivo'] = None
    df.loc[0, 'Objetivo'] = 35
    # Agregar la columna 'Objetivo Global' con el valor 4000 porque son 1000 ventas por mes y nos devuelve 4 meses, en la primera fila.
    df['Objetivo Global'] = 1000
    

    # Crear la nueva columna "Total Aprobadas" con valores 1 o 0
    df['Total Aprobadas'] = df['Estado de Solicitud'].apply(lambda x: 1 if x in ['APROBADA',
                                                                                 'FFTH INSTALADA CON ÉXITO']  else 0)

    # Crear columna de 'Aprobadas Fibra'

    df['Aprobadas Fibra'] = df['Estado de Solicitud'].apply(lambda x: 1 if x == 'FFTH INSTALADA CON ÉXITO' else 0)

    # Crear la columna 'Total Fibras' usando apply y una función lambda
    df['Total Fibras'] = df['Estado de Solicitud'].apply(lambda x: 1 if x in ['FFTH INSTALADA CON ÉXITO',
                                                                          'FTTH RECLAMADA',
                                                                          'FTTH CANCELADA',
                                                                          'FTTH POSTE EN N',
                                                                          'FFTH ACTIVADA EN ESPERA DE INSTALACION',
                                                                          'FFTH PARA ACTIVAR'] else 0)


    # Escribir los datos en la hoja "Sheet1" del archivo "ReporteDiario.xlsx"
    df.to_excel(reporte_diario_path, index=False, sheet_name="Sheet1", na_rep="")


    print("Reporte creado con exito!")
    
    # Eliminar los dos reportes
    os.remove(r"C:\Users\Magellan Banyuls\Downloads\Reporte1.csv")
    os.remove(r"C:\Users\Magellan Banyuls\Downloads\Reporte2.csv")

    # Cargar el archivo Excel nuevamente
    df_reporte = pd.read_excel(r"C:\Users\Magellan Banyuls\Downloads\ReporteDiario.xlsx")

    # Reemplazar las celdas que contienen "" por NaN
    df_reporte = df_reporte.replace("", pd.NA)

    # Guardar el DataFrame con las celdas modificadas en el mismo archivo Excel
    df_reporte.to_excel(r"C:\Users\Magellan Banyuls\Downloads\ReporteDiario.xlsx", index=False, sheet_name="Sheet1")

    print("Eliminadas las comillas...espero")
    # Correo p/Nico
    enviar_correo_adjunto(destinatario, asunto, cuerpo, archivo_adjunto, remitente, contraseña)
    # Correo p/Diego
    enviar_correo_adjunto(destinatario1, asunto, cuerpo1, archivo_adjunto, remitente, contraseña)
    # Correo para Martin
    enviar_correo_adjunto(destinatario2, asunto, cuerpo2, archivo_adjunto, remitente, contraseña)
    # Correo a la Noe
    enviar_correo_adjunto(destinatario4, asunto, cuerpo4, archivo_adjunto, remitente, contraseña)
    # Correo Lean
    enviar_correo_adjunto(destinatario3, asunto, cuerpo3, archivo_adjunto, remitente, contraseña)
    # Correo Deivid
    enviar_correo_adjunto(destinatario5, asunto, cuerpo5, archivo_adjunto, remitente, contraseña)
    



def enviar_correo_adjunto(destinatario, asunto, cuerpo, archivo_adjunto, remitente, contraseña):
    # Configurar el servidor SMTP de Gmail
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Crear el mensaje y establecer los campos "De", "Para" y "Asunto"
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto

    # Agregar el cuerpo del mensaje
    msg.attach(MIMEText(cuerpo, 'plain'))

    # Obtener solo el nombre del archivo sin la ruta completa
    nombre_archivo = os.path.basename(archivo_adjunto)

    # Adjuntar el archivo Excel
    with open(archivo_adjunto, 'rb') as adjunto:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(adjunto.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {nombre_archivo}')
        msg.attach(part)

    # Iniciar sesión en el servidor SMTP y enviar el correo electrónico
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(remitente, contraseña)
        server.sendmail(remitente, destinatario, msg.as_string())
        server.quit()
        print("Correo enviado correctamente.")
    except Exception as e:
        print("Error al enviar el correo:", e)

# Configurar los datos del correo
remitente = 'Aqui va mi mail'
contraseña = 'Aqui va mi contraseña de aplicación'
destinatario = 'nicolas.visaguirre@gmail.com'
asunto = 'Reporte para Tablero de Ventas'
cuerpo = 'Hola Nico! Adjunto ReporteDiario de ventas actualizado a esta hora.'
destinatario1='diegorbl7@gmail.com'
cuerpo1= 'Hola Diego! Adjunto ReporteDiario de ventas actualizado a esta hora.'
destinatario2= 'martinriossoto05@gmail.com'
cuerpo2= 'Hola Martin! Te adjunto ReporteDiario para actualizar!'
destinatario4='noelialopezfrites3010@gmail.com'
cuerpo4='Hola Noe! Te adjunto ReporteDiario para actualizar!'
destinatario3='leandroavila93@gmail.com'
cuerpo3='Hola Lean! Te adjunto ReporteDiario para actualizar!'
destinatario5='carpenter65@hotmail.com'
cuerpo5='Hola Deivid! Adjunto ReporteDiario de ventas actualizado a esta hora.'
# Ruta del archivo Excel a adjuntar
archivo_adjunto = r"C:\Users\Magellan Banyuls\Downloads\ReporteDiario.xlsx"

# Configurar desde que fecha a que fecha
fecha_desde= "01/06/2023"
fecha_hasta= "31/09/2023"


if __name__ == "__main__":
    main()