# openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

def creandoExcel():

    #Iniciamos las liberias para crear el excel
    book = Workbook()
    #Activamos a edicion de columnas 
    sheet = book.active

    #Selecionamos la columans que a las que vamos a agregar info
    sheet['A1'] = "Id"
    sheet['B1'] = "Username"
    sheet['C1'] = "Password"
    sheet['D1'] = "Fecha de creacion"

    #Damos un poco de diseño a las columnas 
    sheet['A1'].font = Font(color='FF0000', bold=True)
    sheet['B1'].font = Font(color='18DD12', bold=True)
    sheet['C1'].font = Font(color='E1239E', bold=True)
    sheet['D1'].font = Font(color='FF0000', bold=True)
    
    #Guaramos el excel
    book.save('bd_login.xlsx')

creandoExcel()