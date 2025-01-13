'''
sačuvati formate
file  picker+export folder
2 modula: dijeljenje excelice +  union all
'''

#from openpyxl import Workbook
import openpyxl
import tkinter as tk
from tkinter import simpledialog
from copy import copy

def main_func():

    # Show the input dialog
    user_input = 'C:/Users/sinis/Documents/Test1.xlsx'
    print ('user_input= '+user_input)



    wb=openpyxl.load_workbook(user_input)
    ws=wb['sheet_a']

    header=list()

    for col in ws.iter_cols(max_row=1):

        #print(col)

        for cell in col:

            #print(cell.value, end=',')

            header.append(cell.value)



    key_col_idx= 0

    for item in header:
        key_col_idx +=1

        if item=='oznaka':
            break      

    keys = set()

    for col in ws.iter_cols(min_row=2,min_col=key_col_idx,max_col=key_col_idx):
        #print(col)

        for cell in col:
            #print(cell.value, end=',')

            if cell.value is not None:
                keys.add(cell.value)    

    curr_row=[]
    

    for key in keys:
        head_col_idx=0
        export_workbook = openpyxl.Workbook()
        export_sheet = export_workbook.active
        
        for col in ws.iter_cols(max_row=1):
        #print(col)

            for  cell in col: 

                    head_col_idx+=1


                    header_cell = export_sheet.cell(row=1, column=head_col_idx, value=cell.value)
            
                    # Copy the cell styles
                    if cell.has_style:
                        header_cell.font = copy(cell.font)
                        header_cell.border = copy(cell.border)
                        header_cell.fill = copy(cell.fill)
                        header_cell.number_format = copy(cell.number_format)
                        header_cell.protection = copy(cell.protection)
                        header_cell.alignment = copy(cell.alignment)

        row_idx=1

        for row in ws.iter_rows(min_row=2):
            #row[0].row -oznacava redak odd prvog elementa u tupleu
            

            if ws.cell(row=row[0].row, column=key_col_idx).value==key:
                #print(key+','+ ws.cell(row=row[0].row, column=key_col_idx).value )
                row_idx+=1
                curr_row=[]
                

                for cell in row:
                    #curr_row.append(cell.value)

                    # Copy the cell value
                    dest_cell = export_sheet.cell(row=row_idx, column=cell.column, value=cell.value)
            
                    # Copy the cell styles
                    if cell.has_style:
                        dest_cell.font = copy(cell.font)
                        dest_cell.border = copy(cell.border)
                        dest_cell.fill = copy(cell.fill)
                        dest_cell.number_format = copy(cell.number_format)
                        dest_cell.protection = copy(cell.protection)
                        dest_cell.alignment = copy(cell.alignment)


                #export_sheet.append(curr_row)
                export_workbook.save(filename="C:/Users/sinis/Documents/"+ key +"_export_test1.xlsx") 

main_func()