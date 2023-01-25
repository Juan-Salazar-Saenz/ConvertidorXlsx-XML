from cmath import e
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
root.geometry("850x300")
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
            hojadeExcel = archivoExcel[nombreHoja.get()]                    #Seleccionamos la hoja de excel
            registros.set(len(hojadeExcel['A'])-1)                          #Contamos el numero de lineas
            archivoXML = escribirXML()                                      #Creamos el archivo de texto .XML
            archivoXML.write('<?xml version="1.0" encoding="UTF-8"?> \n')   #Escribirmos la primera linea de archivo XML
            textoXML = construccionXML(hojadeExcel)                         #Enviamos datos a escribir
            archivoXML.write(str(textoXML))                                 #Escribirmos XML
            archivoExcel.close                                              #Cerramos el archivo
        except e:
            print(e)
            registros.set("Hoja inexistente,¡ intente de nuevo !")        
    except e:
        print(e)
        registros.set("Seleccione Excel,¡ intente de nuevo !")      

#Escrbiir todo el excel separado por comas
#           for fila in hojadeExcel:
#                datos = [celda.value for celda in fila]                
#                textoXML = "" 
#                for i in range(len(datos)):   
#                    if datos[i] == None:
#                        textoXML += "Vacio;"       
#                    else:
#                        textoXML += str(datos[i]) + ";"                    
#                
#                #archivoXML.write(textoXML + " \n")        #Escribimos archivo            
  
                
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

def construccionXML(hojadeExcel):
    
    xml = cxml.Element('garantias')

    op = cxml.SubElement(xml, "op")

    cxml.SubElement(op, "t").text=str("I")
    cxml.SubElement(op, "tg").text=str(registros.get())
    cxml.SubElement(op, "hash").text=str("3925")
    
    i=0                                     #Eliminamos el encabezado de la hoja
    for fila in hojadeExcel:
        if (i>0):   
            gcl = cxml.SubElement(xml, "gcl")         
            datos = [celda.value for celda in fila]  
                    
            ddor = cxml.SubElement(gcl, "ddor")

            cxml.SubElement(ddor, "cci").text=str("1")
            cxml.SubElement(ddor, "ni").text=str(datos[1])
            nombre = str(datos[2]) + str(datos[3])
            cxml.SubElement(ddor, "pn").text=str(nombre)
            cxml.SubElement(ddor, "pa").text=str(datos[4])
            cxml.SubElement(ddor, "sa").text=str(datos[5])
            cxml.SubElement(ddor, "pais").text=str(datos[6])
            cxml.SubElement(ddor, "dpto").text=str(datos[7])
            cxml.SubElement(ddor, "mun").text=str(datos[8])
            cxml.SubElement(ddor, "dir").text=str(datos[9])
            cxml.SubElement(ddor, "email").text=str(datos[10])
            cxml.SubElement(ddor, "tel").text=str(datos[11])
            cxml.SubElement(ddor, "cel").text=str(datos[12])
            cxml.SubElement(ddor, "tdc").text=str("0")
            cxml.SubElement(ddor, "ins").text=str("false")
            cxml.SubElement(ddor, "gen").text=str("1")
            cxml.SubElement(ddor, "tddor").text=str(datos[13])
        
            acdor = cxml.SubElement(gcl, "acdor")

            cxml.SubElement(acdor, "cci").text="2"
            cxml.SubElement(acdor, "ni").text=str(datos[14])
            cxml.SubElement(acdor, "dv").text=str(datos[15])
            cxml.SubElement(acdor, "rs").text=str(datos[16])
            cxml.SubElement(acdor, "pais").text=str(datos[17])
            cxml.SubElement(acdor, "dpto").text=str(datos[18])
            cxml.SubElement(acdor, "mun").text=str(datos[19])
            cxml.SubElement(acdor, "dir").text=str(datos[20])
            cxml.SubElement(acdor, "email").text=str(datos[21])
            cxml.SubElement(acdor, "tel").text=str(datos[22])
            cxml.SubElement(acdor, "cel").text=str(datos[23])
            cxml.SubElement(acdor, "ppal").text=str("true")
            cxml.SubElement(acdor, "ppar").text=str("100")

            descbien = cxml.SubElement(gcl, "descbien").text=str(datos[24])
            prad = cxml.SubElement(gcl, "prad").text=str("true")
            bienes = cxml.SubElement(gcl, "bienes")

            cxml.SubElement(bienes, "ctb").text=str("1")
            cxml.SubElement(bienes, "marca").text=str(datos[26])
            cxml.SubElement(bienes, "srial").text=str(datos[27])
            cxml.SubElement(bienes, "mdlo").text=str(datos[28])
            cxml.SubElement(bienes, "placa").text=str(datos[29])
            cxml.SubElement(bienes, "desc").text=str(datos[30])
            cxml.SubElement(bienes, "fabric").text=str(datos[31])

            monto = cxml.SubElement(gcl, "monto").text=str(datos[32])
            vdef = cxml.SubElement(gcl, "vdef").text=str(datos[33])
            ffin = cxml.SubElement(gcl, "ffin").text=str(datos[34])
            ctg = cxml.SubElement(gcl, "ctg").text="1"
            cm = cxml.SubElement(gcl, "cm").text=str(datos[35])
        i += 1
    
    cxml.indent(xml)                                  #identamos el XML creado
    return str(cxml.tostring(xml, encoding='unicode'))     #Imprimimos el XML

tkinter.Button (root, text=" * Convertir * ", command=convertir).place(x = 100, y = 130) 

tkinter.Label(root, text="Ubicacion de salida").place(x = 10, y = 160) 

tkinter.Label(root, state=DISABLED, textvariable=salida).place(x = 130, y = 160) 

root.mainloop()