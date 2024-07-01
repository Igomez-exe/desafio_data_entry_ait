from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
import os
from pathlib import Path
import datetime
import requests

class Automatizacion:
    def __init__ (self, driver_path):
        self.option = webdriver.ChromeOptions()

        self.driver = webdriver.Chrome(options=self.option)
        self.driver.get("https://desafiodataentryfront.vercel.app/")
        time.sleep(4)

        self.ruta_descargas = self.obtener_ruta_descargas()
        
        self.descargar_primer_lista()
        self.driver.get("https://desafiodataentryfront.vercel.app/")

        self.descargar_segunda_lista()
        self.driver.get("https://desafiodataentryfront.vercel.app/")
        time.sleep(5)

        self.driver.quit()
        self.procesar_primer_lista()
        time.sleep(5)
        self.procesar_segunda_lista()
        time.sleep(3)
        fecha_hoy = datetime.datetime.today().strftime('%Y-%m-%d')

        url = "https://desafio.somosait.com/api/upload/"
        archivo_final_uno = "./AutoFix Repuestos_"+fecha_hoy+".xlsx"
        files = {'file' : open(archivo_final_uno,'rb')}

        response = requests.post(url, files=files)      

        if response.status_code == 200:
            # Éxito: Obtener el link de Google Drive para acceder al archivo subido
            link_drive = response.json().get('link')
            print(f'Archivo subido exitosamente a Google Drive. Link: {link_drive}')
        elif response.status_code == 400:
            # Error: La API devuelve un error 400 si falta alguna columna requerida
            print('Error: Missing required columns.')
        else:
            # Otro código de estado indica un problema diferente
            print(f'Error al subir el archivo. Código de estado: {response.status_code}')


    #AutoRepuestos Express
    def descargar_primer_lista(self):
        auto_respuestos_express = self.driver.find_element(By.XPATH, "/html/body/div/div/div/section/div/div[1]/div[2]/button").click()
        time.sleep(3)

    #AutoFix
    def descargar_segunda_lista(self):
        self.driver.find_element(By.XPATH,"/html/body/div/div/div/section/div/div[2]/div[2]/button").click()
        time.sleep(3)

        usuario = self.driver.find_element(By.NAME, "username")
        contraseña = self.driver.find_element(By.NAME, "password")
        usuario.send_keys("desafiodataentry")
        contraseña.send_keys("desafiodataentrypass")
        contraseña.send_keys(Keys.RETURN)

        time.sleep(3)
        self.marcar_palomitas()
        time.sleep(3)
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[2]/button").click()
        
        #esto podria llegar a traer problemas.
        time.sleep(60)

    #Mundo RepCar
    def descargar_tercer_lista(self):
        self.driver.find_element(By.XPATH, "/html/body/div/div/div/section/div/div[3]/div[2]/button").click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, "/html/body/div/div/div/main/div/div/div[1]/section[1]/div/div/div/div[2]/button/a").click()
        time.sleep(3)



    def procesar_primer_lista(self):
        primer_archivo = 'AutoRepuestos Express.xlsx' 
        primer_lista = self.ruta_descargas / primer_archivo
        #print(primer_lista)

        try:
            df = pd.read_excel(primer_lista, sheet_name="AJ_AP_1719248653832.xlsx")
            fila_inicio = 11
            fila_fin = 15050

            valores_codigos = df.loc[fila_inicio-1:fila_fin-1, 'AutoRepuestos Express'].tolist()
            valores_descripcion = df.loc[fila_inicio-1:fila_fin-1,'Unnamed: 1'].tolist()
            valores_marca = df.loc[fila_inicio-1:fila_fin-1,'Unnamed: 7'].tolist()
            valores_precios = df.loc[fila_inicio-1:fila_fin-1, 'Unnamed: 2'].tolist()

        except FileNotFoundError:
            print(f"El archivo no se encontró en la ruta especificada: {primer_lista}")
        except Exception as e:
            print(f"Ocurrió un error al leerel archivo excel: {e}")

        fecha_hoy = datetime.datetime.today().strftime('%Y-%m-%d')
        nombre_salida = f"{"AutoRepuestos Express"}_{fecha_hoy}.xlsx"
        ruta_salida = os.path.join('./', nombre_salida)
        print(f"Archivo procesado y guardado en: {ruta_salida}") 

        nuevo_excel = (
            pd.DataFrame({'CODIGO':valores_codigos, 'DESCRIPCION':valores_descripcion, 'MARCA':valores_marca, 'PRECIO':valores_precios})
            )
        
        archivo_excel = './AutoRepuestos Express_'+fecha_hoy+'.xlsx'
        nuevo_excel.to_excel(archivo_excel, index=False)
        
    def procesar_segunda_lista(self):
        segundo_archivo = "AutoFix Repuestos.xlsx"
        segunda_lista = self.ruta_descargas / segundo_archivo

        try:

            df_fiat = pd.read_excel(segunda_lista, sheet_name="FI")
            fila_inicio = 1
            fila_fin = 11909
            valores_codigos_fiat = df_fiat.loc[fila_inicio-1:fila_fin-1, 'CODIGO'].tolist()
            valores_descrpcion_fiat = df_fiat.loc[fila_inicio-1:fila_fin-1,'DESCR'].tolist()
            valores_marca_fiat = ["FIAT"]*11908
            valores_precios_fiat = df_fiat.loc[fila_inicio-1:fila_fin-1, 'PRECIO'].tolist()
            
            df_chevrolet = pd.read_excel(segunda_lista, sheet_name="CH")
            fila_inicio_ch = 1
            fila_fin_ch = 10234
            valores_codigos_ch = df_chevrolet.loc[fila_inicio_ch-1:fila_fin_ch, 'CODIGO'].tolist()
            valores_descripcion_ch = df_chevrolet.loc[fila_inicio_ch-1:fila_fin_ch-1, 'DESCR'].tolist()
            valores_marca_ch = ["CHEVROLET"]*10233
            valores_precios_ch = df_chevrolet.loc[fila_inicio_ch-1:fila_fin_ch-1, 'PRECIO'].tolist()

            df_vw = pd.read_excel(segunda_lista, sheet_name="VW")
            fila_inicio_vw = 1
            fila_fin_vw = 13030
            valores_codigos_vw = df_vw.loc[fila_inicio_vw-1:fila_fin_vw-1, 'CODIGO'].tolist()
            valores_descripcion_vw = df_vw.loc[fila_inicio_vw-1:fila_fin_vw-1, 'DESCR'].tolist()
            valores_marca_vw = ["VOLKSWAGEN"]*13029
            valores_precios_vw = df_vw.loc[fila_inicio_vw-1:fila_fin_vw-1, 'PRECIO'].tolist()  

            df_ci = pd.read_excel(segunda_lista, sheet_name="CI")
            fila_inicio_ci = 1
            fila_fin_ci = 5157
            valores_codigos_ci = df_ci.loc[fila_inicio_ci-1:fila_fin_ci-1, 'CODIGO'].tolist()
            valores_descripcion_ci = df_ci.loc[fila_inicio_ci-1:fila_fin_ci-1, 'DESCR'].tolist()
            valores_marca_ci = ["CITROEN"]*5156
            valores_precios_ci = df_ci.loc[fila_inicio_ci-1:fila_fin_ci-1, 'PRECIO'].tolist()

            df_ty = pd.read_excel(segunda_lista, sheet_name="TY")
            fila_inicio_ty = 1
            fila_fin_ty = 2641
            valores_codigos_ty = df_ty.loc[fila_inicio_ty-1:fila_fin_ty-1, 'CODIGO'].tolist()
            valores_descripcion_ty = df_ty.loc[fila_inicio_ty-1:fila_fin_ty-1, 'DESCR'].tolist()
            valores_marca_ty = ["TOYOTA"]*2640
            valores_precios_ty = df_ty.loc[fila_inicio_ty-1:fila_fin_ty-1, 'PRECIO'].tolist()

            df_fo = pd.read_excel(segunda_lista, sheet_name="FO")
            fila_inicio_fo = 1
            fila_fin_fo = 9817
            valores_codigos_fo = df_fo.loc[fila_inicio_fo-1:fila_fin_fo-1, 'CODIGO'].tolist()
            valores_descripcion_fo = df_fo.loc[fila_inicio_fo-1:fila_fin_fo-1, 'DESCR'].tolist()
            valores_marca_fo = ["FORD"]*9816
            valores_precios_fo = df_fo.loc[fila_inicio_fo-1:fila_fin_fo-1, 'PRECIO'].tolist()

            df_rn = pd.read_excel(segunda_lista, sheet_name="RN")
            fila_inicio_rn = 1
            fila_fin_rn = 19931
            valores_codigos_rn = df_rn.loc[fila_inicio_rn-1:fila_fin_rn-1, 'CODIGO'].tolist()
            valores_descripcion_rn = df_rn.loc[fila_inicio_rn-1:fila_fin_rn-1, 'DESCR'].tolist()
            valores_marca_rn = ["RENAULT"]*19930
            valores_precios_rn = df_rn.loc[fila_inicio_rn-1:fila_fin_rn-1, 'PRECIO'].tolist()
            


        except FileNotFoundError:
            print(f"El archivo no se encontró en la ruta especificada: {segunda_lista}")
        except Exception as e:
            print(f"Ocurrió un error al leerel archivo excel: {e}")

        fecha_hoy = datetime.datetime.today().strftime('%Y-%m-%d')
        nombre_salida = f"{"AutoFix Repuestos"}_{fecha_hoy}.xlsx"
        ruta_salida = os.path.join('./', nombre_salida)
        print(f"Archivo procesado y guardado en: {ruta_salida}") 

        nuevo_excel = (
            pd.DataFrame({'CODIGO':valores_codigos_fiat + valores_codigos_ch + valores_codigos_vw + valores_codigos_ci + valores_codigos_ty + valores_codigos_fo + valores_codigos_rn, 
                          'DESCRIPCION':valores_descrpcion_fiat + valores_descripcion_ch + valores_descripcion_vw + valores_descripcion_ci + valores_descripcion_ty + valores_descripcion_fo + valores_descripcion_rn,
                          'MARCA':valores_marca_fiat + valores_marca_ch + valores_marca_vw + valores_marca_ci + valores_marca_ty + valores_marca_fo + valores_marca_rn,
                          'PRECIO':valores_precios_fiat + valores_precios_ch + valores_precios_vw + valores_precios_ci + valores_precios_ty + valores_precios_fo + valores_precios_rn})
            )
        
        archivo_excel = './AutoFix Repuestos'+'_'+fecha_hoy+'.xlsx'
        nuevo_excel.to_excel(archivo_excel, index=False)






    def obtener_ruta_descargas(self):        
        if os.name == 'nt':
            return Path(os.getenv('USERPROFILE')) / 'Downloads'
        
        #en macOs / Linux es  return Path.home() / 'Downloads' 

    #Para AutoFix (esto se puede mejorar)
    def marcar_palomitas(self):
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[1]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[2]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[3]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[4]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[5]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[6]/input").click()
        self.driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/main/div[2]/form/div/div[1]/fieldset/div/div[7]/input").click()


if __name__ == "__main__":
    bot = Automatizacion("chromedriver.exe")