import pyodbc
import sys
import time
import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from six.moves import urllib
from pathlib import Path
from datetime import date
from datetime import datetime



ExisteERR = False
timeout = False
usr_peoplesoft =''
pas_peoplesoft = ''
Link_Descarga = ''
nombreInterfaz = ''


direccion_servidor = 'localhost'
nombre_bd = 'stage'
nombre_usuario = 'sa'
password = 'etrora135'


#webdrive_path ="D:/robotizacion/chromedriver.exe"

webdrive_path ="D:/Procesos de Carga/Descarga Interfaces PeopleSoft/chromedriver.exe"


ConfTimeOut= 900
contador =0


def FinProceso():
    print("***********************************************************************")
    print("Final  del proceso :" + str(datetime.now()))
    print("***********************************************************************")

def IniProceso():
    print("***********************************************************************")
    print("Inicio del proceso :" + str(datetime.now()))
    print("***********************************************************************")

def ActualizaInterfazDescargada(vnombrefile,vnombreInterfaz,vfecha_busqueda):
	try:
		conexion = pyodbc.connect('DRIVER={SQL Server};SERVER=' + direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
	except Exception as e:
		# Atrapar error
		print('Ocurrio un error al conectar a SQL Server: ', e)
	try:
		with conexion.cursor() as cursor:
			vfecha = str(datetime.now().strftime('%Y-%b -%d %H:%M'))
			SQLCommand = "UPDATE PROC_INTERFAZ_DETALLE SET FECHA_DESCARGA_INTERFAZ='"+vfecha +"', [NOMBRE_INTERFAZ_DESCARGA] ='"+ vnombrefile + "'  WHERE interfaz ='" + vnombreInterfaz+ "' and cast(fecha_busqueda_casilla as date) ='" + vfecha_busqueda + "'"
			print(SQLCommand)
			cursor.execute(SQLCommand) 
	except Exception as e:
		print("Ocurrio un error al consultar: ", e)
	finally:
		conexion.close()
		
def ActualizaInterfazTimeout(vnombreInterfaz,vfecha_busqueda):
	try:
		conexion = pyodbc.connect('DRIVER={SQL Server};SERVER=' + direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
	except Exception as e:
		# Atrapar error
		print('Ocurrio un error al conectar a SQL Server: ', e)
	try:
		with conexion.cursor() as cursor:
			vfecha = str(datetime.now().strftime('%Y-%b -%d %H:%M'))
			SQLCommand = "UPDATE PROC_INTERFAZ_DETALLE SET FECHA_DESCARGA_INTERFAZ='"+vfecha +"', [ESTADO_DESCARGA] ='TIMEOUT'  WHERE interfaz ='" + vnombreInterfaz+ "' and cast(fecha_busqueda_casilla as date) ='" + vfecha_busqueda + "'"
			print(SQLCommand)
			cursor.execute(SQLCommand) 
	except Exception as e:
		print("Ocurrio un error al consultar: ", e)
	finally:
		conexion.close()

fecha =date.today()
strfecha = format(str(fecha.year) + str(fecha.month) + str(fecha.day))
a=0
#Genera la fecha para el log



#intenta recuperar los argumentos
try:
    if len(sys.argv) >= 2:
        nombreInterfaz  = sys.argv[1]
        Link_Descarga   = sys.argv[2]
        ruta_descarga   = sys.argv[3]
        fecha_busqueda  = sys.argv[4]
        ruta_para_log   = sys.argv[5]
        usr_peoplesoft  = sys.argv[6]
        pas_peoplesoft  = sys.argv[7]
		
        NombreLog = ruta_para_log + '\\' + nombreInterfaz + '_' + strfecha + '.log'
        # se setea en nombre del archivo de log
        sys.stdout = open(NombreLog, 'w')
        IniProceso()
        for a in range(a,len(sys.argv)):
             print("argv[" + str(a) +"] = " + sys.argv[a] )
    
except:
        sys.stdout = open('DescargaInterfaznula.log', 'w')
        IniProceso()
        print('Error en control de argumentos no se definieron los 7 parametros')
        for a in range(a,len(sys.argv)):
              print("argv[" + str(a) +"] = " + sys.argv[a] )
        raise ValueError('No ha especificado un nombre de interfaz.')
        FinProceso()
        raise

#valida el link de descarga

if ruta_descarga == '':
	print("No se ha establecido una ruta de descarga para la interfaz")
	FinProceso()
	raise ValueError('Error en conexion al server')
else:
	print("Ruta de Descarga:" + ruta_descarga)

#exit()
	   
#Se inicia la descarga de las interfaces

preferences = {"download.default_directory": ruta_descarga ,
               "directory_upgrade": True,
               "safebrowsing.enabled": True }



options = webdriver.ChromeOptions()
options.add_argument('--headless')            # Habilita o desahabilita la interfaz grafica
options.add_argument('--ignore-certificate-errors')
options.add_experimental_option("prefs", preferences)
options.add_argument('--disable-extensions')
options.add_argument('--log-level=3')



try:
    driver  = webdriver.Chrome(webdrive_path,chrome_options=options)
except Exception as e:
    print("Problemas al cargar webdrive" + str(e))
    raise ValueError('Problemas al cargar webdrive' + str(e))
print(Link_Descarga)
driver.get(Link_Descarga)
linkbtn = "/html/body/table[1]/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td[1]/table[2]/tbody/tr[4]/td[4]/input"


txt_user = driver.find_element_by_id('userid')
txt_pass = driver.find_element_by_id('pwd')
btn_log  =driver.find_element_by_xpath(linkbtn)


#Carga las variables recogidas desde la base de datos
txt_user.send_keys(usr_peoplesoft )
txt_pass.send_keys(pas_peoplesoft )
btn_log.click()

try:
    div_errorlog = driver.find_element_by_id('login_error')
    style_div_error = div_errorlog.get_attribute('style')
except:
    style_div_error = 'No existe el div'
    
if style_div_error == 'visibility: visible;':
    print('Error en las credenciales utilizadas para el sitio')
   
else:
    #aqui se recarga la pagina y podemos encontrar nuevos objetos
    try:
        lnk_descarga = driver.find_element_by_id('URL$1')
    except:
        print("El link del reporte ha expirado")
        driver.quit()
        raise
    
    rutadescarga = lnk_descarga.get_attribute('href');
    vector = rutadescarga.split("/")
    nombrefile = vector[6]
    DOWNLOAD_PATH =ruta_descarga + "\\" + nombrefile

    fileObj = Path(DOWNLOAD_PATH)
    #verificamos si existe el archivo a descargar
    #si existe lo borramos
    if fileObj.is_file():
        os.remove(DOWNLOAD_PATH)
        print("Se borro el arhivo " + nombrefile)

    #se hace el click para descargar el arhivo  
    print("Inicia descarga del archivo")  
    lnk_descarga.click();

    #se ejecuta el ciclo hasta que el archivo se genere
    while True:
        if fileObj.is_file():
            break;
        else:  
            time.sleep(2)
            contador=contador+2
        if contador==ConfTimeOut:
             timeout = True
             break;
        
    if timeout == True:
        print("Timeout complido: no se descargo el archivo")
        ActualizaInterfazTimeout(nombreInterfaz,fecha_busqueda)
        FinProceso()
        driver.quit()
        #raise ValueError('timeout cumplido')
    else:    
        print("Archivo descargado: " + DOWNLOAD_PATH)
        ActualizaInterfazDescargada(DOWNLOAD_PATH,nombreInterfaz,fecha_busqueda)
#se finaliza la ejecucion del python
driver.quit()
FinProceso()
	