# -*- coding: utf-8 -*-
"""
Created on Fri Dec 17 13:59:19 2021

@author: OABI
"""

import win32com.client
import os, time


directory = os.getcwd()
files = os.listdir(directory)
LAlist = []
TRMlist = []


# first the loop diagams in a list
for file in files:
    LoopAtrribs_dict = {}
    if '.DWG' in file:
        if '-DL-' in file:
            print('LOOPDIAGRAM',file) 
            filename = "/" + file
            acad = win32com.client.Dispatch("Autocad.Application")
            doc = acad.Documents.Open(directory+filename)
            acad.Visible = False
            for entity in acad.ActiveDocument.PaperSpace:
                    name = entity.ObjectName
                    if name == 'AcDbBlockReference':
                        HasAttributes = entity.HasAttributes
                        if HasAttributes:
                            print(file)
                            LoopAtrribs_dict['DOCUMENT_NUMBER'] = file
                            for attrib in entity.GetAttributes():
                                LoopAtrribs_dict[attrib.TagString] = attrib.TextString
#                                print(attrib.TagString,attrib.TextString)
#                                LoopAtrribs_dict[attrib.TagString]=attrib.TextString                              
            LAlist.append(LoopAtrribs_dict)
            doc.close(False)
            time.sleep(2)
acad.Application.Quit()
#print(LAlist)
path = os.getcwd() + "\LPAttributes.xlsx"
if os.path.exists(path):
    os.remove(path)
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Xcel.Visible = False
Wb= Xcel.Workbooks.Add()
Ws = Wb.ActiveSheet
column = 1
for key in LAlist[0]:
    Xcel.ActiveSheet.Cells(1,column).Value = key
    column += 1
row = 2
column =1
for dictionary in LAlist:
    for key in dictionary:
        Xcel.ActiveSheet.Cells(row,column).Value = dictionary[key]
        column += 1
    row += 1
    column = 1
Wb.ActiveSheet.Range("A:EE").Columns.AutoFit()    
Wb.SaveAs(path)
Xcel.Application.Quit()
       



#For TermDiagrams

for file in files:
    TermAtrribs_dict = {}
    if '.DWG' in file:
        if '-DT-' in file or '-DR-' in file or '-DG-' in file:
            print('TERMINATIONDIAGRAM',file) 
            filename = "/" + file
            acad = win32com.client.Dispatch("Autocad.Application")
            doc = acad.Documents.Open(directory+filename)
            acad.Visible = False
            for entity in acad.ActiveDocument.PaperSpace:
                    name = entity.ObjectName
                    if name == 'AcDbBlockReference':
                        HasAttributes = entity.HasAttributes
                        if HasAttributes:
#                            print(file)
                            TermAtrribs_dict['DOCUMENT_NUMBER'] = file
                            for attrib in entity.GetAttributes():
                                TermAtrribs_dict[attrib.TagString] = attrib.TextString
#                                print(attrib.TagString,attrib.TextString)
#                                LoopAtrribs_dict[attrib.TagString]=attrib.TextString                              
            TRMlist.append(TermAtrribs_dict)
            doc.close(False)
            time.sleep(2)
acad.Application.Quit()
#print(LAlist)
path = os.getcwd() + "\TERMAttributes.xlsx"
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Xcel.Visible = False
Wb= Xcel.Workbooks.Open(path)
Ws = Wb.ActiveSheet
column = 1
for key in TRMlist[0]:
    Xcel.ActiveSheet.Cells(1,column).Value = key
    column += 1
row = 2
column =1
for dictionary in TRMlist:
    for key in dictionary:
        Xcel.ActiveSheet.Cells(row,column).Value = dictionary[key]
        column += 1
    row += 1
    column = 1
Wb.ActiveSheet.Range("A:EE").Columns.AutoFit()    
Wb.Close(SaveChanges=True)
Xcel.Application.Quit()     
#    

    
print("COMPLETED")               