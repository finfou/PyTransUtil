#!/usr/bin/env python3
# -*- coding:utf-8 -*-
from openpyxl import Workbook
import xml.etree.ElementTree as etree


if __name__ == '__main__':
    tree = etree.parse('strings_call_recorder.xml')
    root = tree.getroot()

    wb = Workbook()
    ws = wb.active

    # WriteHead(wb)
    ws['A1']='Source'
    ws['B1']='Key'
    ws['C1']='Notes'
    ws['D1']='English'
    ws['E1']='Chinese'
    ws['F1']='Hindi'
    ws['H1']='Indonesia'

    # Write entries
    cnt=2
    for child in root:
        print(child.attrib['name'] + ' = ' + child.text)
        ws['A'+cnt]=


    # wb = Workbook()
    #
    # # grab the active worksheet
    # ws = wb.active
    #
    # # Data can be assigned directly to cells
    # ws['A1'] = 42
    #
    # # Rows can also be appended
    # ws.append([1, 2, 3])
    #
    # # Python types will automatically be converted
    # import datetime
    # ws['A2'] = datetime.datetime.now()
    #
    # # Save the file
    # wb.save("sample.xlsx")