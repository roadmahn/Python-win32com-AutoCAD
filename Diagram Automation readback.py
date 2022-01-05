# -*- coding: utf-8 -*-
"""
Created on Fri Dec 17 13:59:19 2021

@author: OABI
"""

import win32com.client
import os, time


directory = r'Y:\1_PROJS\1-1807-0\Engineering - ElecInstr\Python Scripts\AS BUILT'
files = os.listdir(directory)
LAlist = []
TRMlist = []
# first the loop diagams in a list

##open Excel file and copy the info into a dict nd append that Dict into  a list

#LoopAttributes excel file
#path = os.getcwd() + "\LPAttributes.xlsx"
#Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
#Xcel.Visible = False
#Wb= Xcel.Workbooks.Open(path)
#Ws = Wb.ActiveSheet
#LoopAtrribs_dict = {}
#LA_dict = {}
#xlToLeft = -4159
#xlUp = -4162
#LastCol = Xcel.ActiveSheet.Cells(1, Xcel.ActiveSheet.Columns.Count).End(xlToLeft).Column
#LastRow = Xcel.ActiveSheet.Cells(Xcel.ActiveSheet.Rows.Count, "A").End(xlUp).Row
#print(LastRow,LastCol)
#for row in range(2,LastRow+1):
#    for column in range(1,LastCol+1): 
#        header = Xcel.ActiveSheet.Cells(1,column).Value
#        attrb = Xcel.ActiveSheet.Cells(row,column).Value
#        LoopAtrribs_dict[header] = attrb
#    for key in LoopAtrribs_dict:
#        if LoopAtrribs_dict[key] is not None:
#            LA_dict[key] = LoopAtrribs_dict[key]
##    print(LA_dict)
#    LAlist.append(LA_dict)
#    LA_dict ={}
#    LoopAtrribs_dict = {}
##    print(LAlist)
#Wb.Close(SaveChanges=True)
#Xcel.Application.Quit()
#time.sleep(1)

##TermAttributes excel file
path = os.getcwd() + "\TERMAttributes.xlsx"
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Xcel.Visible = False
Wb= Xcel.Workbooks.Open(path)
Ws = Wb.ActiveSheet
TermAtrribs_dict = {}
TA_dict ={}
xlToLeft = -4159
xlUp = -4162
LastCol = Xcel.ActiveSheet.Cells(1, Xcel.ActiveSheet.Columns.Count).End(xlToLeft).Column
LastRow = Xcel.ActiveSheet.Cells(Xcel.ActiveSheet.Rows.Count, "A").End(xlUp).Row
print(LastRow,LastCol) 
for row in range(2,LastRow+1):
    for column in range(1,LastCol+1): 
        header = Xcel.ActiveSheet.Cells(1,column).Value
        attrb = Xcel.ActiveSheet.Cells(row,column).Value
        TermAtrribs_dict[header] = attrb
    for key in TermAtrribs_dict:
        if TermAtrribs_dict[key] is not None:
            TA_dict[key] = TermAtrribs_dict[key]
#    print(TA_dict)
    TRMlist.append(TA_dict)
    TA_dict ={}
    TermAtrribs_dict = {}
Wb.Close(SaveChanges=True)
Xcel.Application.Quit()
time.sleep(1)
 
            
##open autocad and read out LPattributes and append Values in Attribute Dict to TextString.   

##LoopDiagrams       
#for file in files:
#    if '.DWG' in file:
#        if '-DL-' in file:
#            print('LOOPDIAGRAM',file) 
#            filename = "/" + file
#            acad = win32com.client.Dispatch("Autocad.Application")
#            doc = acad.Documents.Open(directory+filename)
#            acad.Visible = False
#            for entity in acad.ActiveDocument.ModelSpace:
#                name = entity.ObjectName
#                if name == 'AcDbPolyline':
#                    if entity.Layer == 'Revision':
#                        entity.Color =  '18'
#                        entity.Update()
#                if name == 'AcDbBlockReference':
#                    HasAttributes = entity.HasAttributes
#                    if HasAttributes:
#                        if entity.Layer == 'Revision':
#                            entity.Color =  '18'
#                            entity.Update()
#                            for attrib in entity.GetAttributes():
##                                print(entity.Layer,attrib.TagString, attrib.TextString)
#                                if attrib.TagString == '1' :
#                                   attrib.TextString = '' 
#                                   attrib.Update()
#            print('LOOPDIAGRAM',file) 
#            for entity in acad.ActiveDocument.PaperSpace:
#                name = entity.ObjectName
#                if name == 'AcDbBlockReference':
#                    HasAttributes = entity.HasAttributes
#                    if HasAttributes:
#                        for attrib in entity.GetAttributes():
##                            print(attrib.TagString, attrib.TextString)
#                            if attrib.TextString == '':
#                                for LA_dict in LAlist:
#                                    if LA_dict['DOCUMENT_NUMBER'] != file:
#                                        continue
#                                    for key in LA_dict:
#                                            if attrib.TagString == key: 
#                                                attrib.TextString = LA_dict[key]
#                                                attrib.Update()
#                                                print(attrib.TagString, attrib.TextString, LA_dict[key])
##                                print(LA_dict)
#            doc.SaveAs(directory+filename)
#            doc.close()
#acad.Application.Quit()
#time.sleep(1)

# Termination diagrams
for file in files:
    if '.DWG' in file:
        if '-DT-' in file or '-DR-' in file or '-DG-' in file:
            print('TERMINATIONDIAGRAM',file) 
            filename = "/" + file
            acad = win32com.client.Dispatch("Autocad.Application")
            doc = acad.Documents.Open(directory+filename)
            acad.Visible = False
            for entity in acad.ActiveDocument.ModelSpace:
                name = entity.ObjectName
                if name == 'AcDbPolyline':
                    if entity.Layer == 'Revision':
                        entity.Color =  '18'
                        entity.Update()
                if name == 'AcDbBlockReference':
                    HasAttributes = entity.HasAttributes
                    if HasAttributes:
                        if entity.Layer == 'Revision':
                            entity.Color =  '18'
                            entity.Update()
                            for attrib in entity.GetAttributes():
#                                print(entity.Layer,attrib.TagString, attrib.TextString)
                                if attrib.TagString == '1' :
                                   attrib.TextString = '' 
                                   attrib.Update()
            print('TERMINATIONDIAGRAM',file) 
            for entity in acad.ActiveDocument.PaperSpace:
                name = entity.ObjectName
                if name == 'AcDbBlockReference':
                    HasAttributes = entity.HasAttributes
                    if HasAttributes:
                        for attrib in entity.GetAttributes():
#                            print(attrib.TagString, attrib.TextString)
                            if attrib.TextString == '':
                                for TA_dict in TRMlist:
                                    if TA_dict['DOCUMENT_NUMBER'] != file:
                                        continue
                                    for key in TA_dict:
                                            if attrib.TagString == key: 
                                                attrib.TextString = TA_dict[key]
                                                attrib.Update()
                                                print(attrib.TagString, attrib.TextString, TA_dict[key])
#                                print(LA_dict)
            doc.SaveAs(directory+filename)
            doc.close()
acad.Application.Quit()
time.sleep(1)

               
print("COMPLETED")