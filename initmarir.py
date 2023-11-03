import PyPDF2
import re
import os
import openpyxl
from xls2xlsx import XLS2XLSX
from openpyxl.styles import PatternFill
import pyautogui
import ctypes  # An included library with Python install.
import requests

banescofiles=[]
mercantilfiles=[]
bncfiles=[]
references=[]
# creating a pdf file object
def processpdf(filename):
    referencespdf=[]
    banktype=None
    pdfFileObj = open(str(filename), 'rb')
    if re.findall('mercantil',str(((filename).lower()))):
        banktype='mercantil'
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    for page in pdfReader.pages:
        if banktype=='mercantil':
            lines=(page.extract_text()).split('\n')
            for line in lines:
                if re.search('[0-9][0-9]+/+[0-9][0-9]+/+[0-9][0-9]',line):
                    reference=str((line.split(' '))[1])
                    if reference.startswith('000000'):
                        print('detectado ref provincial')
                    if reference.startswith('0000'):
                        print('detectado ref provincial')
                    # print(reference)
                    referencespdf.append(str(reference))
                
    referencespdf=set(referencespdf)
    return referencespdf
def processxls(filename):
    referencesexcel=[]
    wb_obj = openpyxl.load_workbook(filename)
    ws = wb_obj.active
    banktype=None
    if re.findall('banesco',str(((filename).lower()))):
        banktype='banesco'
    if re.findall('bnc',str(((filename).lower()))):
        banktype='bnc'
    for row in ws.iter_rows():
        for cell in row:
            if banktype=='banesco':
                refletter='B'
            if banktype=='bnc':
                refletter='M'
            try:
                if cell.column_letter==refletter:
                    if re.search('\d+',str((cell.value))):
                        cellvalue=str(cell.value)
                        if cellvalue.startswith('000000'):
                            print('detectado provincial')
                            cellvalue=cellvalue[5:]
                        if cellvalue.startswith('0000'):
                            print('detectado provincial')
                        # print(cell.value)
                        
                        referencesexcel.append(str(cellvalue))
            except AttributeError:
                pass
            
    return set(referencesexcel)
        
def processfactura(filename,referecelist):
    def colorcell(color):
        if color=='red':
            color = openpyxl.styles.colors.Color(rgb='00FF0000')
        if color=='green':
            color = openpyxl.styles.colors.Color(rgb='00008000')
        my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=color)
        return my_fill
    referencesonfacts=[]
    validsreferences=[]
    invalidscolumns=[]
    validscolumns=[]
    wb_obj = openpyxl.load_workbook(filename)
    ws = wb_obj.active
    for row in ws.iter_rows():
        for cell in row:
            if str((cell.value)) =='Referencia':
                reflocation=str((cell.coordinate)[0])
            try:
                if cell.column_letter==reflocation:
                    try:
                        if re.search('\d+',str((cell.value))):
                            extractedref=re.search('\d+',str((cell.value))).group()
                            if extractedref.startswith('000000'):
                                print('detectado ref provincial 6 ceros')
                                extractedref=extractedref[6:]
                            if extractedref.startswith('0000'):
                                print('detectado ref 4 ceros')
                                extractedref=extractedref[4:]
                            if extractedref.startswith('000'):
                                print('detectado ref 3 ceros')
                                extractedref=extractedref[3:]
                            if extractedref.startswith('00'):
                                print('detectado ref 2 ceros')
                                extractedref=extractedref[2:]
                            if extractedref.startswith('0'):
                                print('detectado ref 1 cero')
                                extractedref=extractedref[1:]
                            # print(extractedref)
                            if extractedref in referecelist:
                                validscolumns.append(cell.coordinate)
                                print('encontre referencia valida')
                            else:
                                invalidscolumns.append(cell.coordinate)
                    except:
                        pass
            except UnboundLocalError:
                pass
    
    for coordinates in invalidscolumns:
        print(coordinates)
        ws[coordinates].fill = colorcell('red')
    for coordinates in validscolumns:
        ws[coordinates].fill = colorcell('green')
        # ws[coordinates].fill = 'greenFill'
    wb_obj.save(filename)  # save the workbook
    wb_obj.close()
    return str(len(validscolumns)), str(len(invalidscolumns))
for file in os.listdir():
    if re.findall('factura',str(((file).lower()))):
        if file.endswith('.xls'):
            x2x = XLS2XLSX(file)
            newfilename=str(file[:-3])+'xlsx'
            os.remove(file)
            wb = x2x.to_xlsx(newfilename)
            facturafilename=(newfilename)
            if re.findall('banesco',str(((file).lower()))):
                banescofiles.append(file)
            if re.findall('mercantil',str(((file).lower()))):
                mercantilfiles.append(file)
            if re.findall('bnc',str(((file).lower()))):
                bncfiles.append(file)
        if file.endswith('.xlsx'):
            facturafilename=file
            if re.findall('banesco',str(((file).lower()))):
                banescofiles.append(file)
            if re.findall('mercantil',str(((file).lower()))):
                mercantilfiles.append(file)
            if re.findall('bnc',str(((file).lower()))):
                bncfiles.append(file)
    else:
        if file.endswith('.pdf'):
            if re.findall('banesco',str(((file).lower()))):
                banescofiles.append(file)
            if re.findall('mercantil',str(((file).lower()))):
                mercantilfiles.append(file)
            if re.findall('bnc',str(((file).lower()))):
                bncfiles.append(file)
            references.extend(processpdf(file))
        if file.endswith('.xls'):
            x2x = XLS2XLSX(file)
            newfilename=str(file[:-3])+'xlsx'
            os.remove(file)
            wb = x2x.to_xlsx(newfilename)
            references.extend(processxls(newfilename))
            if re.findall('banesco',str(((file).lower()))):
                banescofiles.append(file)
            if re.findall('mercantil',str(((file).lower()))):
                mercantilfiles.append(file)
            if re.findall('bnc',str(((file).lower()))):
                bncfiles.append(file)
        if file.endswith('.xlsx'):
            references.extend(processxls(file))
            if re.findall('banesco',str(((file).lower()))):
                banescofiles.append(file)
            if re.findall('mercantil',str(((file).lower()))):
                mercantilfiles.append(file)
            if re.findall('bnc',str(((file).lower()))):
                bncfiles.append(file)
if len(bncfiles)==0:
    response=pyautogui.confirm(text='no se han encontrado los movimientos de bnc\ndesea continuar ?', title='error', buttons=['Si', 'No'])
    if response =='No':
        exit()
if len(banescofiles)==0:
    response=pyautogui.confirm(text='no se han encontrado los movimientos de banesco\ndesea continuar ?', title='error', buttons=['Si', 'No'])
    if response =='No':
        exit()
if len(mercantilfiles)==0:
    response=pyautogui.confirm(text='no se han encontrado los movimientos de mercantil\ndesea continuar ?', title='error', buttons=['Si', 'No'])
    if response =='No':
        exit()
if not facturafilename is None:
    valids,invalids=processfactura(facturafilename,set(references))
    response=pyautogui.confirm(text='se han encontrado '+valids+' referencias en la factura\nse han encontrado '+invalids+' invalidas ', title='error', buttons=['Aceptar'])
