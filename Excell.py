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

    Wbmod = load_workbook('modelo.xlsx')
    WSmod = Wbmod['Factura Simpli']

    WSmod['D5'].value = Fecha
    WSmod['D7'].value = NumFact
    WSmod['A11'].value = Nombre
    WSmod['A12'].value = Ip
    WSmod['A13'].value = Pais
    WSmod['A17'].value = Descripcion
    #Cambiamos el . por la , para que excell pueda trabajar con el numero
    WSmod['D17'].value = Euros.replace('.' , ',')

    #Comprobamos que el IVA sea distinto a 0
    if(IVA == '0.0' ):
            WSmod['E23'].value = '0'
  

    Wbmod.save('Facturas/'+str(NumFact) + '.xlsx')

    print( 
        # WSmod['D5'].value +
    #WSmod['D7'].value +
    #WSmod['A11'].value +
    #WSmod['A12'].value +
    #WSmod['A13'].value +
    #WSmod['A17'].value +
    #WSmod['D17'].value +
    #str(WSmod['E23'].value)



       # ("Factura numero " + NumFact + " Fecha de Factura " + Fecha  
       # + " Cantidad pagada "  + Euros + " Euros " + "Nombre cliente "+ Nombre + " del pais "+ Pais +
        # " buscado por la Ip "+ Ip + " Se le grava un IVA de " + IVA + " Compro el video "+ Descripcion)
        )





#Para recorrer todo el archivo sacando la informacion que necesitamos.

for i in range(2,122):
    Printarventa(str(i))