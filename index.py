from tkinter import *                #Codigo para interfaces en python
from tkinter import filedialog       #Codigo para importar desde windows en ventana
import tkinter
import openpyxl                      #Codigo para leer archivos excel

#configuracion de la venta
root = Tk()
root.geometry("450x300")
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
            hojadeExcel = archivoExcel[nombreHoja.get()] 
            registros.set(len(hojadeExcel['A'])-1)
        except:
            registros.set("Hoja inexistente,¡ intente de nuevo !")        
    except:
        registros.set("Seleccione Excel,¡ intente de nuevo !")        
    
        #for fila in hojadeExcel:
        #    print([celda.value for celda in fila])
        

tkinter.Button (root, text="Convertir", command=convertir).place(x = 100, y = 130) 

root.mainloop()