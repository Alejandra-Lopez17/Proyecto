# openpyxl
from openpyxl import *
from openpyxl import Workbook
from datetime import date

def creandoUsuario():
    #Iniciamos el excel
    book = load_workbook('bd_login.xlsx')
    
    #Recuperamos la cantidad de filas de nuestro codigo
    max_row=book.active.max_row
    
    print(max_row)
    
    #Datos de usuario
    userName = input("Ingresa nombre de usuario: ")
    password = input("Ingresa contraseña(mayor de 5 digitos)(Con lestras y numeros): ")
    confirPassword = input("confirme su contraseña: ")
    
    
    #Confirmamos la contraseña
    if password == confirPassword and len(password) > 5:
        #Agregar id
        sheet = book.active
        sheet[f"A{max_row+1}"]= max_row
        
        #Agregar username
        sheet = book.active
        sheet[f"B{max_row+1}"]= userName
        
        #Agregar password
        sheet = book.active
        sheet[f"C{max_row+1}"]= password
        
        #Agregar fecha de creacion 
        sheet = book.active
        sheet[f"D{max_row+1}"]= date.today()
        
        #Guardamos el excel
        print("Usuario Registrado")
        book.save('bd_login.xlsx')
    else:
        print("Contraseña no valida")
        
creandoUsuario()