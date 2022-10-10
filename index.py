from dataclasses import dataclass
from platform import architecture
from tkinter import *                #Codigo para interfaces en python
from tkinter import filedialog       #Codigo para importar desde windows en ventana
import tkinter
import openpyxl                      #Codigo para leer archivos excel
from io import open                  #Codigo para escribir archivo en txt

#configuracion de la venta
root = Tk()
root.geometry("650x300")
root.title("Convertidor de .xlsx a XML BBVA Colombia - Loans")
root.resizable(0,0)
root.iconbitmap("bbva.ico")
root.config(relief="groove")
root.config(cursor="hand2")
root.config(bd="45")

#titulo
tkinter.Label(root, text="Convertir de .xlsx a XML", bg="blue").pack(anchor=CENTER) 

#variables de ficheros
ubicacion = StringVar()
registros = StringVar()
nombreHoja = StringVar()
salida = StringVar()

def abreFichero():
    #Buscamos la ruta del archivo
    ubicacion.set(filedialog.askopenfilename(title="Abrir", initialdir="c:", filetypes=[("Ficheros de Excel","*.xlsx")]))
    
    
    #sheetHoja = archivoExcel.sheet_by_name("hoja1")
    #registros.set(sheetHoja.nrows)
    
tkinter.Button (root, text="Seleccionar fichero", command=abreFichero).place(x = 10, y = 30) 

tkinter.Label(root, state=DISABLED, textvariable=ubicacion).place(x = 130, y = 30) 

tkinter.Label(root, text="Nombre de la hoja").place(x = 10, y = 70) 

tkinter.Entry(root, textvariable=nombreHoja).place(x = 130, y = 70) 

tkinter.Label(root, text="Cantidad de registros").place(x = 10, y = 100) 

tkinter.Label(root, state=DISABLED, textvariable=registros).place(x = 130, y = 100) 

def convertir():    
    try:
        #Se paramos el archivo en carpetas
        ruta = ubicacion.get().split(sep="/")
        #Recorremos el archivo para cambiar el / por //
        ubicacionExcel = ""
        for texto in ruta:    
            ubicacionExcel += texto + "//"
        #reconstruimos la ruta
        ubicacionExcel=ubicacionExcel[0:len(ubicacionExcel)-2]
        #abrimos el archivo excel
        archivoExcel = openpyxl.load_workbook(ubicacionExcel, data_only=True)
        #Seleecionamos la hoja
        try:
            hojadeExcel = archivoExcel[nombreHoja.get()]  #Seleccionamos la hoja de excel
            registros.set(len(hojadeExcel['A'])-1)        #Contamos el numero de lineas
            archivoXML = escribirXML()                    #Creamos el archivo de texto .XML
            
            for fila in hojadeExcel:
                datos = [celda.value for celda in fila]
                print(datos)
                textoXML = "" 
                for i in range(len(datos)):   
                    if datos[i] == None:
                        textoXML += "Vacio;"       
                    else:
                        textoXML += str(datos[i]) + ";"                    
                
                archivoXML.write(textoXML + " \n")        #Escribimos archivo            
            
            archivoExcel.close                            #Cerramos el archivo
        except:
            registros.set("Hoja inexistente,¡ intente de nuevo !")        
    except:
        registros.set("Seleccione Excel,¡ intente de nuevo !")        
                
def escribirXML():

    #Se paramos el archivo en carpetas
    ruta = ubicacion.get().split(sep="/")
    #Recorremos el archivo para cambiar el / por //
    ubicacionXml = ""
    ubicacionXml1 = ""
    
    for i in range(len(ruta)-1):    
        ubicacionXml += ruta[i] + "//"
        ubicacionXml1 += ruta[i] + "/"
    #reconstruimos la ruta
    nombretxt = ruta[len(ruta)-1].split(sep=".")
    ubicacionXml += nombretxt[0] + ".XML"
    ubicacionXml1 += nombretxt[0] + ".XML"
    
    try:
        salida.set(ubicacionXml1)
        archivoTexto = open(ubicacionXml,"w")
        return archivoTexto
    except:
        salida.set("Archivo existente,¡ intente de nuevo !")

tkinter.Button (root, text=" * Convertir * ", command=convertir).place(x = 100, y = 130) 

tkinter.Label(root, text="Ubicacion de salida").place(x = 10, y = 160) 

tkinter.Label(root, state=DISABLED, textvariable=salida).place(x = 130, y = 160) 

root.mainloop()