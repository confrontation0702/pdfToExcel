# -*- coding: utf-8 -*-
import os
from win32com import client

def eTP():
    xlApp = client.DispatchEx("Excel.Application")
    path = os.path.abspath('./out/')
    file_names = os.listdir("./out/")
    for file in file_names:
        if file.endswith(".xlsx"):
            print("-------------")
            print(file)
            excel = os.path.join(path, file)
            pdf = excel.replace("xlsx","pdf")
            books = xlApp.Workbooks.Open(excel, False)
            books.ExportAsFixedFormat(0, pdf)
    xlApp.Quit()

#eTP()
