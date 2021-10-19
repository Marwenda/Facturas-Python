#Marwenda version 1.00
#Facturas de Listado a modelo, NO se adjunta en el repositorio ningun tipo de informacion sensible
#Variables autoexplicativas(la mayoria)


#Creamos una execpcion por si falla la apertura de la libreria.
try:
    from openpyxl.utils.cell import get_column_letter
    from openpyxl import Workbook, load_workbook
except ImportError:
    # Que hacer si el módulo no se puede importar
    print("Módulo no instalado")

#Inicializamos el libro contable que nos interesa
wb = load_workbook('registro.xlsx')
#Inicializamos solo la hoja del libro que nos interesa
ws = wb['Facturas simpli']



#Creamos una funcion para que solo pasandole un numero me de los valores que esperamos de la venta
def Printarventa(NumeroFila):
   #Guardamos todos los valores que nos interesan en variables que usaremos despues.
    NumFact = str(ws['A' + NumeroFila].value)
    Fecha = str(ws['B' + NumeroFila].value)
    Euros = str(ws['H' + NumeroFila].value)
    Ip = str(ws['M' + NumeroFila].value)
    Pais = str(ws['N' + NumeroFila].value)
    IVA = str(ws['I' + NumeroFila].value)
    Nombre = str(ws['E' + NumeroFila].value)+" "+str(ws['F' + NumeroFila].value)
    Descripcion= str(ws['G' + NumeroFila].value)

#Abrimos el libro del modelo de factura, y escogemos la hoja de factura simplificada
    Wbmod = load_workbook('modelo.xlsx')
    WSmod = Wbmod['Factura Simpli']
#Rellenamos los valores que nos interesan del registro.
    WSmod['D5'].value = Fecha
    WSmod['D7'].value = NumFact
    WSmod['A11'].value = Nombre
    WSmod['A12'].value = Ip
    WSmod['A13'].value = Pais
    WSmod['A17'].value = Descripcion
    #Cambiamos el . por la , para que excell pueda trabajar con el numero
    WSmod['D17'].value = Euros.replace('.' , ',')

    #Comprobamos que el IVA sea distinto a 0, si no por defecto es 21%
    if(IVA == '0.0' ):
            WSmod['E23'].value = '0'
  
#Guardamos cada factura con su propio libro para subirlo a Drive o transformarlo en PDF a posteriori.

    Wbmod.save('Facturas/'+str(NumFact) + '.xlsx')

    # print( # ("Factura numero " + NumFact + " Fecha de Factura " + Fecha  
       # + " Cantidad pagada "  + Euros + " Euros " + "Nombre cliente "+ Nombre + " del pais "+ Pais +
        # " buscado por la Ip "+ Ip + " Se le grava un IVA de " + IVA + " Compro el video "+ Descripcion))





#Para recorrer todo el archivo sacando la informacion que necesitamos.
#Usamos el rango de 2,122 para sacar las primeras 121 Facturas simplificadas.
#Funciona perfectamente.

for i in range(2,122):
    Printarventa(str(i))

#TODO Necesitamos interfaz grafica, y un panel de opciones para configurar tanto la celda que queramos para obtener
# los datos como para que celda rellenar de forma sencilla.    