import base
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import *
import xlsxwriter

root= tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global df
    import_file_path = filedialog.askopenfilename()
    patho=str(import_file_path).split('.')
    print(len(patho))
    pat=''
    ran=len(patho)-1
    for z in range(ran):
        if z==ran-1:
            pat = pat + patho[z]
        else:
            pat=pat+patho[z]+"/"
    print(pat)
    file=pd.read_excel(import_file_path)
    workbook = xlsxwriter.Workbook(pat+"-2.xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    worksheet.write('A1', 'Tel', bold)
    worksheet.write('B1', 'nir')
    worksheet.write('C1', 'serie')
    worksheet.write('D1', 'compania')
    worksheet.write('E1', 'tipo')
    worksheet.write('F1', 'municipio')

    canvas1.create_text(100,75,text="Termine", font=18,fill='red')
    cad=""
    print("Filas:"+str(file.size))
    row = 1
    col = 0
    for y in range(file.size):
        db=base.conexion()
        mycur=db.cursor()
        num=str(file.at[y,'tel'])
        if num=="nan":
            continue
        nir=num[0:3]
        serie=num[3:6]
        res=num[6:10]
        sql="select nir, serie, razon_social, tipo_red, municipio from plan2 where nir="+nir
        sql=sql+" and serie="+serie +" and NUMERACION_INICIAL<"+res +" and NUMERACION_FINAL>"+res
        mycur.execute(sql)
        result=mycur.fetchall()
        for x in result:
            print(result[0][1])
            print(row)
            worksheet.write_string(row,col,str(num))
            worksheet.write_string(row,col+1,str(result[0][0]))
            worksheet.write_string(row, col+2, str(result[0][1]))
            worksheet.write_string(row, col+3, str(result[0][2]))
            worksheet.write_string(row, col+4, str(result[0][3]))
            worksheet.write_string(row, col+5, str(result[0][4]))
            cad=cad+str(num)+str(x)+" "
            row =row+1
            #print(num+"-")
            #print(x)
    workbook.close()

browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)
root.mainloop()