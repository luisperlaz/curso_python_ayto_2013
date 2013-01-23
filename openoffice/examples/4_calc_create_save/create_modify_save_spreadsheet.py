# -.- coding: utf-8 -.-

import uno
import os


def get_context(host="localhost", port="2002"):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    connurl ="uno:socket,host={},port={};urp;StarOffice.ComponentContext"
    return resolver.resolve(connurl.format(host, port))


ctx = get_context()
serviceManager = ctx.ServiceManager
desktop = serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

# Hasta aquí obtención del contexto, servicemanager, y desktop

# Creación de documento
url = "private:factory/scalc"
doc = desktop.loadComponentFromURL(url, "_blank", 0, ()) #SpreadSheetDocument


### Operaciones básicas con hojas y celdas

# Insertamos hoja
sheets = doc.getSheets() #XSpreadSheets
sheets.insertNewByName("Primera hoja", 0)
sheet = sheets.getByIndex(0)

# Obtenemos una celda y le damos valor
xCell = sheet.getCellByPosition(0, 0)
xCell.setValue(123)

##### Guardado

doc.storeAsURL("file:///tmp/testsheet.ods", ())

##### Exportado

prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
prop.Name = "FilterName"
prop.Value = "MS Excel 97"
doc.storeAsURL("file:///tmp/testsheet.xls", (prop,)); 

