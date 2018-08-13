import pandas
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from datetime import datetime

tiposComp = {
"factura a" : "001",
"factura b" : "006",
"factura c" : "011",
"ticket-factura a" : "081",
"ticket-factura b" : "082",
"n/c a" : "003",
"n/c b" : "008",
"n/d a" : "002",
"n/d b" : "007"
}

tiposArch = [
("Excel XML", ".xlsx"),
("Excel", ".xls")
]

def abreArchivo():
    ruta = askopenfilename(title ="Seleccione archivo",
    filetypes = tiposArch).replace("\n","")
    eRuta.delete(0, END)
    eRuta.insert(0, ruta)
    planilla = pandas.ExcelFile(ruta)
    hojas = planilla.sheet_names
    listaHojas.delete(0, END)
    for hoja in hojas:
        listaHojas.insert(END, hoja)

def tipoComp(c, l):
    tipo = "{} {}".format(c, l)
    if tipo.lower() in tiposComp.keys():
        return tiposComp[tipo.lower()]
    else:
        return "DESCONOCIDO!! AVISAR!!"

def convImp(nro):
    if 0 < nro < 999999999:
        return str(int(nro * 100)).rjust(15, "0")
    else:
        return str(int(0)).rjust(15, "0")

def lineaAlic(li, neto, iva, al):
    #tipo comprobantes, pto, nro
    retorno = li[8:36]
    #documento vendedor, identificación vendedor
    retorno += li[52:74]
    #neto gravado, alicuota, impuesto
    retorno += convImp(neto) + al + convImp(iva)

    return retorno

def nan0(n):
    if 0 < n < 99999999:
        return n
    else:
        return 0

def nan1(n):
    if not type(n) == str:
        return 1
    else:
        return n

def exportar():
    planilla = pandas.ExcelFile(eRuta.get())
    datos=planilla.parse(listaHojas.get(ACTIVE))
    datos.columns=['A', 'B', 'C', 'D','E', 'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
    salida = asksaveasfilename(defaultextension=".txt", filetypes=[("text", "*.txt")] )
    salidaAlic = salida.replace(".txt", "_alic.txt")
    with open(salida, "w") as archivo:
        with open(salidaAlic, "w") as archivoAlic:
            for i in range(0, len(datos.index)-1):
                #exportación de comprobantes
                if type(datos.loc[i,'A']) != str and type(datos.loc[i,'A']) != float:
                    #fecha
                    linea = datetime.strftime(datos.loc[i,'A'], "%Y%m%d")
                    #tipo
                    linea += tipoComp(str(datos.loc[i,'B']), str(datos.loc[i,'C']))
                    #pto, nro
                    linea += "{}{}".format(str(nan1(datos.loc[i,'D'])).rjust(5,"0"),
                    datos.loc[i,'E'].rjust(20,"0"))
                    # despacho importación
                    linea += "".rjust(16, " ")
                    #documento vendedor
                    linea += "80".rjust(2, " ")
                    #identificación vendedor
                    linea += datos.loc[i,'H'].replace("-","").rjust(20, "0")
                    #nombre vendedor
                    if len(datos.loc[i,'G']) < 31:
                        linea += datos.loc[i,'G'].replace("Ñ", "N").ljust(30, " ")
                    else:
                        linea += datos.loc[i,'G'][:30].replace("Ñ", "N")
                    #total de la operacion
                    linea += convImp(datos.loc[i, 'T'])
                    #conceptos que no integran el NG
                    linea += convImp(0)
                    #operaciones exentas
                    linea += convImp(0)
                    #percepcione a cta de IVA
                    linea += convImp(datos.loc[i, 'Q'])
                    #percep imp nac
                    linea += convImp(0)
                    #percep IIBB
                    linea += convImp(nan0(datos.loc[i, 'R']) + nan0(datos.loc[i, 'S']))
                    #percep imp munic
                    linea += convImp(0)
                    #impuestos internos
                    linea += convImp(datos.loc[i, 'P'])
                    #cod moneda
                    linea += "PES"
                    #tipo de cambio
                    linea += "0001000000"
                    #cant de alic de IVA m n o
                    j = 0
                    if datos.loc[i, 'M'] > 0:
                        j +=1
                    if datos.loc[i, 'N'] > 0:
                        j += 1
                    if datos.loc[i, 'O'] > 0:
                        j += 1
                    linea += str(j)
                    #cod de operacion
                    linea += " "
                    #cred fiscal computable
                    linea += convImp(nan0(datos.loc[i, 'M']) + nan0(datos.loc[i, 'N'])
                     + nan0(datos.loc[i, 'O']))
                    #otros tributos
                    linea += convImp(0)
                    #cuit emisor/corredor
                    linea += "".rjust(11, "0")
                    #nombre emisor/corredor
                    linea += "".ljust(30, " ")
                    #IVA comision
                    linea += convImp(0)

                    #exportacion de alicuota
                    if -99999999 < datos.loc[i, 'J'] < 99999999:
                        archivoAlic.write(lineaAlic(linea, datos.loc[i, 'J'],
                        datos.loc[i, 'M'], "0004") + "\r\n")
                    if -99999999 < datos.loc[i, 'K'] < 99999999:
                        archivoAlic.write(lineaAlic(linea, datos.loc[i, 'K'],
                        datos.loc[i, 'N'], "0005") + "\r\n")
                    if -99999999 < datos.loc[i, 'L'] < 99999999:
                        archivoAlic.write(lineaAlic(linea, datos.loc[i, 'L'],
                        datos.loc[i, 'O'], "0006") + "\r\n")

                    archivo.write(linea + "\r\n")

#ventana principal
ventana = Tk(className="Excel a COMPRAS")

varLabRuta = StringVar()
varLabRuta.set("Ruta al archivo Excel:")
labRuta = Label(ventana, textvariable=varLabRuta)
labRuta.grid(row="0", column="0")

ruta = StringVar()
eRuta = Entry(ventana, textvariable=ruta)
eRuta.grid(row="1",column="0", columnspan="7")

bRuta = Button(ventana, text="...", command=abreArchivo)
bRuta.grid(row="1", column="8")

varLabHojas=StringVar()
varLabHojas.set("Seleccione la hoja de COMPRAS:")
labHojas=Label(ventana, textvariable=varLabHojas)
labHojas.grid(row="3", column="0")

listaHojas = Listbox(ventana)
listaHojas.grid(row="4",column="0", columnspan="8")

dExportar = Button(ventana, text="Exportar", command=exportar)
dExportar.grid(row="5", column="0")

ventana.mainloop()
