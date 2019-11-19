# -*- coding: utf-8 -*-
import glob
import pdfplumber
import xlwings as xw
import excelToPdf
import repairNumber
import json
from shutil import copyfile

#写入部件行
def write_unit_cells(ws,units):
    for index, v in enumerate(units):
        print(index)
        row = 9+index
        ws.range('B%d' % (row)).value = v[1]
        ws.range('D%d' % (row)).value = v[5]
    #模板默认100行部件模板数据，删除空行
    ws.range('%d:108' % (len(units)+9)).api.EntireRow.Delete()

def write_titles(ws,titles):
    #解析维修设备 维修编号
    splitList = [' 日期 Date: ',
                 '\n推荐的维修等级 Service Level: ',
                 ' 客户名称 Client: ',
                 ':\n维修号 Build No: ',
                 '客户编号 Client No',
                 '\n序列号 SN: ',
                 '工号 Job No.: ',
                 '维修设备 Equip: ']
    cellList = ['F7','C7','F6','C6','F5','C5','F4','C4']
    tmpList = titles[0]
    titleList = []
    for split in splitList:
        t = tmpList.split(split)
        titleList.append(t[1])
        tmpList = t[0]
    #获取维修编号    
    customer_id = titleList[4]
    repairNo = repairNumber.creatNo(customer_id)
    titleList[6] = repairNo
    d = romanNoToChinese(titleList[1])
    print('---------------------')
    print(d)
    titleList[1] = d
    #写入excel
    for index ,t in enumerate(titleList):
        ws.range(cellList[index]).value = t
#罗马数字转中文数字
def romanNoToChinese(str):
    levels = []
    with open("./config/level.json",'rb') as ls:
        dicts = json.load(ls)
        levels = dicts['levels']
    nList = sorted(levels, key=lambda l: int(l['arab']), reverse = True)
    print(nList)
    for l in nList:
        if l['roman'] in str:
            print(l['roman'])
            nStr = str.replace(l['roman'],l['chinese'])
            print(nStr)
            return nStr


def analyzePDF():
    unitList = []
    titleList = []
    pdfs = glob.glob('./resource/*.pdf')
    for p in pdfs:
        pdf_name = p.split('\\')[1]
        pdf_name = pdf_name.split('.pdf')[0]
        copyfile("./resource/templates.xlsx", "./out/%s.xlsx"%pdf_name)
        app = xw.App(visible=False, add_book=False)
        workbook =  app.books.open("./out/%s.xlsx"%pdf_name)
        workbook_sh = workbook.sheets["维修估算单（竖版）"]
        unitList = []
        titleList = []
        pdf = pdfplumber.open(p)
        for page in pdf.pages:
            for table in page.extract_tables():
                for index, row in enumerate(table):
                    print(row)
                    if index == 1:
                        titleList = row
                    try:
                        if int(row[0]) and row[1] != "":
                            unitList.append(row)
                    except:
                        titleList.append(row)
        write_titles(workbook_sh,titleList)
        write_unit_cells(workbook_sh,unitList)
    

        pdf.close()
        workbook.save()
        workbook.close()
        app.quit()
    excelToPdf.eTP()
def main():
    analyzePDF()

if __name__ == "__main__":  
    main()
    
