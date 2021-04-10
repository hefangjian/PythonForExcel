import os
from win32com import client

def DocToPDF(doc_name, pdf_name):
    word = client.DispatchEx("Word.Application")
    if os.path.exists(pdf_name):
        os.remove(pdf_name)
    worddoc = word.Documents.Open(doc_name, ReadOnly=1)
    worddoc.SaveAs(pdf_name, FileFormat=17)
    worddoc.Close()
    word.Quit()

#注意是全路径 abspath
# print(os.path.abspath('.'))
#DocToPDF(os.path.join(os.path.abspath('.'),'letter.docx'),os.path.join(os.path.abspath('.'),'abbbc.pdf'))

def ExcelToPDF(Excel_name, pdf_name):
    Excel = client.DispatchEx("Excel.Application")
    if os.path.exists(pdf_name):
        os.remove(pdf_name)
    excel = Excel.Workbooks.Open(Excel_name, ReadOnly=1)
    excel.ExportAsFixedFormat(0, pdf_name)
    excel.Close()
    Excel.Quit()
#ExcelToPDF(os.path.join(os.path.abspath('.'),'VIPLIST.xlsx'),os.path.join(os.path.abspath('.'),'VIP.pdf'))


def pptTOPDF(PPT_name,pdf_name):
    PPT = client.DispatchEx('PowerPoint.Application')
    ppt = PPT.Presentations.Open(PPT_name)
    ppt.ExportAsFixedFormat(pdf_name, 2, PrintRange=None)
    ppt.Close()
    PPT.Quit()
#pptTOPDF(os.path.join(os.path.abspath('..'), 'myFurniture.pptx'), os.path.join(os.path.abspath('..'),  'newFurniture.pdf'))

import pandas as pd
split_excel_name_head = '考试'
split_excel_name_tail = '信息汇总.xlsx'
xlsx_name = '/Users/65106/Desktop/firstexcel.xlsx'
#xlsx_name = 'firstexcel.xlsx'
#用来筛选的列名
filter_column_name = '姓名'
#将该列去重后保存为list
df = pd.read_excel(xlsx_name)
#print(df)
school_names = df[filter_column_name].unique().tolist()
print(school_names)
#获取所有sheet名
df = pd.ExcelFile(xlsx_name)
sheet_names = df.sheet_names
print(sheet_names)
sheet_not_filter_names = sheet_names[1:4]
print(sheet_not_filter_names)
#city_name_to_list = []
for school_name in school_names:
    city_excel_name = split_excel_name_head + str(school_name) + split_excel_name_tail
    writer = pd.ExcelWriter(city_excel_name)
    city_name_to_list = []
    city_name_to_list.append(school_name)
    for sheet_name in sheet_names:
        tmp_df = pd.read_excel(xlsx_name, sheet_name=sheet_name)

        if sheet_name not in sheet_not_filter_names:
           tmp_sheet = tmp_df[tmp_df[filter_column_name].isin(city_name_to_list)]
        else:
           tmp_sheet = tmp_df

        tmp_sheet.to_excel(excel_writer=writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        writer.save()
        writer.close()
        ExcelToPDF(os.path.join(os.path.abspath('.'), city_excel_name), os.path.join(os.path.abspath('.'), split_excel_name_head + str(school_name)+'.pdf'))