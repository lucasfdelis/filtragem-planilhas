import csv
import xlsxwriter
from tkinter import filedialog
from tkinter import *

root = Tk()
root.withdraw()
folder_selected = filedialog.askopenfilename()

while(folder_selected == ''):
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askopenfilename()

file = open(folder_selected, encoding='utf-8')
tabela = csv.reader(file, delimiter='|')

folder_selected = folder_selected.split('.')
folder_selected[1] = '.xlsx'
folder_selected = ''.join(folder_selected)
print(folder_selected)
workbook = xlsxwriter.Workbook(folder_selected)
sheet1 = workbook.add_worksheet()

style = workbook.add_format({'bold': True})
sheet1.write(0, 0, 'CPF',style)
sheet1.write(0, 1, 'NOME',style)
sheet1.write(0, 2, 'CONTATO 1',style)   
sheet1.write(0, 3, 'CONTATO 2',style)
sheet1.write(0, 4, 'CONTATO 3',style)
sheet1.write(0, 5, 'ENDEREÇO',style)
sheet1.write(0, 6, 'PONTO DE REFERENCIA',style)
sheet1.write(0, 7, 'VELOCIDADE',style)
sheet1.write(0, 8, 'FIMAGENDAMENTO',style)
sheet1.write(0, 9, 'OBSERVAÇÕES',style)
k = 1

for l in tabela:
    vcdivisao = l[80]
    operadordivisao = l[48]
    if(l[1]!='Entregue ao técnico'):
        if(l[11]=='INSTALAÇÃO BL E VOIP'):
            if(l[67]=='PB' or l[67]=='AL' or l[67]=='BA' or l[67]=='MG' or l[67] =='PE' or l[67]=='SE'):
                if(l[79]=='VAREJO'):
                    if(vcdivisao[0:3].isdigit()==True):
                        inteirodivisao = vcdivisao[0:3]
                        inteirodivisao = int(inteirodivisao)
                        if(int(inteirodivisao>=200) and vcdivisao[3:8]==' MBPS'):
                            operadordivisao = operadordivisao[0:2]
                            if(operadordivisao == 'BC' or operadordivisao == 'CC'):
                                if(l[16].isdigit()==True):
                                    sheet1.write(k, 0, int(l[16]))
                                else:
                                    sheet1.write(k, 0, l[16])
                                sheet1.write(k, 1, l[43])
                                if(l[16].isdigit()==True):
                                    sheet1.write(k, 2, int(l[17]))
                                else:
                                    sheet1.write(k, 2, l[17])
                                if(l[16].isdigit()==True):
                                    sheet1.write(k, 3, int(l[18]))
                                else:
                                    sheet1.write(k, 3, l[18])
                                if(l[16].isdigit()==True):
                                    sheet1.write(k, 4, int(l[19]))
                                else:
                                    sheet1.write(k, 4, l[19])
                                sheet1.write(k, 5, l[96])
                                sheet1.write(k, 6, l[97])
                                sheet1.write(k, 7, l[80])
                                sheet1.write(k, 8, l[3])
                                k = k+1

workbook.close()
