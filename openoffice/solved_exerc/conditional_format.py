# -.- coding: utf-8 -.-

import uno
import os

def conditional_format():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    serviceManager = ctx.ServiceManager
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheets = doc.getSheets() #XSpreadSheets
    sheet = sheets.getByIndex(0)

    ### Aquí empieza el ejercicio de traducción del ejemplo de formateado condicional

    xCellRange = sheet.getCellRangeByName("A1:B10");
    xEntries = xCellRange.getPropertyValue("ConditionalFormat");

    cond1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    cond2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    cond3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")

    cond1.Name = "Operator";
    cond1.Value = uno.getConstantByName("com.sun.star.sheet.ConditionOperator.GREATER")
    cond2.Name = "Formula"
    cond2.Value = "1"
    cond3.Name = "StyleName"
    cond3.Value = "Heading"
    xEntries.addNew((cond1, cond2, cond3))
    xCellRange.setPropertyValue("ConditionalFormat", xEntries); 
