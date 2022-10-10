from dataclasses import dataclass
from platform import architecture
from tkinter import *                #Codigo para interfaces en python
from tkinter import filedialog       #Codigo para importar desde windows en ventana
import tkinter
import openpyxl                      #Codigo para leer archivos excel
from io import open                  #Codigo para escribir archivo en txt
import xml.etree.ElementTree as cxml #Codigo para escribir archivo en XML

    
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
            archivoXML.write('<?xml version="1.0" encoding="UTF-8"?> \n')
            for fila in hojadeExcel:
                datos = [celda.value for celda in fila]                
                textoXML = "" 
                for i in range(len(datos)):   
                    if datos[i] == None:
                        textoXML += "Vacio;"       
                    else:
                        textoXML += str(datos[i]) + ";"                    
                
                #archivoXML.write(textoXML + " \n")        #Escribimos archivo            

            textoXML = construccionXML()  
            archivoXML.write(textoXML)
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

def construccionXML():
    xml = cxml.Element('garantias')

    op = cxml.SubElement(xml, "op")

    cxml.SubElement(op, "t").text="I"
    cxml.SubElement(op, "tg").text="1"
    cxml.SubElement(op, "hash").text="3925"

    glc = cxml.SubElement(xml, "glc")
    ddor = cxml.SubElement(glc, "ddor")
    
    cxml.SubElement(ddor, "cci").text="1"
    cxml.SubElement(ddor, "ni").text="5022743"
    cxml.SubElement(ddor, "pn").text="ABELARDO"
    cxml.SubElement(ddor, "pa").text="BARRETO"
    cxml.SubElement(ddor, "sa").text="GUILLEN"
    cxml.SubElement(ddor, "pais").text="CO"
    cxml.SubElement(ddor, "dpto").text="8"    
    cxml.SubElement(ddor, "mun").text="08001"
    cxml.SubElement(ddor, "dir").text="CR 9 A 17 19"
    cxml.SubElement(ddor, "email").text="abelardobarretoguillen@hotmail.com"
    cxml.SubElement(ddor, "tel").text="3719150"
    cxml.SubElement(ddor, "cel").text="3114097178"
    cxml.SubElement(ddor, "tdc").text="0"
    cxml.SubElement(ddor, "ins").text="false"
    cxml.SubElement(ddor, "gen").text="1"
    cxml.SubElement(ddor, "tddor").text="d"

    acdor = cxml.SubElement(glc, "acdor")

    cxml.SubElement(acdor, "cci").text="6"
    cxml.SubElement(acdor, "ni").text="860003020"
    cxml.SubElement(acdor, "dv").text="1"
    cxml.SubElement(acdor, "rs").text="BANCO BILBAO VIZCAYA ARGENTARIA COLOMBIA S.A.BBVA COLOMBIA"
    cxml.SubElement(acdor, "pais").text="CO"
    cxml.SubElement(acdor, "dpto").text="11"
    cxml.SubElement(acdor, "mun").text="11001"
    cxml.SubElement(acdor, "dir").text="KR 9 72-21 PI '7'"
    cxml.SubElement(acdor, "email").text="Garantiasmobiliarias@confecamaras.org.co"
    cxml.SubElement(acdor, "tel").text="3471600"
    cxml.SubElement(acdor, "cel").text="0313471600"
    cxml.SubElement(acdor, "ppal").text="true"
    cxml.SubElement(acdor, "ppar").text="100"

    descbien = cxml.SubElement(glc, "descbien").text="Vehiculo"
    prad = cxml.SubElement(glc, "prad").text="true"
    bienes = cxml.SubElement(glc, "bienes")

    cxml.SubElement(bienes, "ctb").text="1"
    cxml.SubElement(bienes, "marca").text="KIA"
    cxml.SubElement(bienes, "srial").text="KNABE511AHT359063"
    cxml.SubElement(bienes, "mdlo").text="2017"
    cxml.SubElement(bienes, "placa").text="WGD259"
    cxml.SubElement(bienes, "desc").text="Vehiculo"
    cxml.SubElement(bienes, "fabric").text="GM Colmotores"

    monto = cxml.SubElement(glc, "monto").text="46350000"
    vdef = cxml.SubElement(glc, "vdef").text="1"
    ffin = cxml.SubElement(glc, "ffin").text="2032-12-30T23:59:59.0000000"
    ctg = cxml.SubElement(glc, "ctg").text="1"
    cm = cxml.SubElement(glc, "cm").text="COP"

    cxml.indent(xml)                                  #identamos el XML creado
    return cxml.tostring(xml, encoding='unicode')     #Imprimimos el XML

tkinter.Button (root, text=" * Convertir * ", command=convertir).place(x = 100, y = 130) 

tkinter.Label(root, text="Ubicacion de salida").place(x = 10, y = 160) 

tkinter.Label(root, state=DISABLED, textvariable=salida).place(x = 130, y = 160) 

root.mainloop()