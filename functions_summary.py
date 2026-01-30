import os
import openpyxl
import copy
from openpyxl.utils import get_column_letter
import re
import pandas as pd


### SUMMARY PACKING LIST ###
#pre processamento dos excels summary packing list

def pre_proc_summary(folder_path, arquivos):
    for file in arquivos:
        wb = openpyxl.load_workbook(f'{folder_path}/{file}')
        ws = wb.active

        max_row = ws.max_row
        for row in range(13, max_row + 1):
            #remover o valor de grid
            ws.cell(row=row, column=9).value = '' #grid

            #selecionar a coluna w.o number e separar em 3 valores, w.o ente, w.o year, wo.number
            cell = ws.cell(row=row, column=3) #coluna w.o number
            val = cell.value
            if not isinstance(val, str):
                continue
            
            val = str(val).strip().replace('  ','')
            m = re.search(r'CH_BM', str(val).strip(), flags=re.IGNORECASE)
            if m:
                parts = val.split(' ')
                print(parts)
                if len(parts) != 3:
                    print(f'O PO {val} não está no formato esperado')
                    print('Formato esperado: CH_BM{espaço em branco}ANO{espaço em branco}NÚMERO')
                    print('Como não estava no formato esperado, as colunas w.o ente, w.o year não foram preenchidas para este PO.')
                else:
                    wo_ente = parts[0]
                    ws.cell(row=row, column=1).value = wo_ente
                    wo_year = parts[1]
                    ws.cell(row=row, column=2).value = wo_year
                    wo_number = parts[2]
                    ws.cell(row=row, column=3).value = wo_number
                
            #selecionar a coluna style name e separar em 2 valores, style name e article number
            cell2 = ws.cell(row=row, column=5) #coluna STYLE NAME
            val2 = cell2.value
            if not isinstance(val2, str):
                continue

            m2 = re.search(r'ART', val2, flags=re.IGNORECASE)
            if m2:
                start = m2.start()
                end = m2.end() # retorna a posição inicial da correspondência
                first = val2[:start].strip()
                ws.cell(row=row, column=5).value = first

                second = val2[end:].strip().replace('.','')
                ws.cell(row=row, column=6).value  = second 
            else:
                print(f'O STYLE NAME {val2} não menciona o valor ART')
                print('Portanto, a coluna ARTICLE ficou vazia') 
                ws.cell(row=row, column=6).value  = ''



        wb.save(f'{folder_path}/{file}')
    return

