# -.- coding: utf-8 -.-

import uno
import os

ctx = XSCRIPTCONTEXT.getComponentContext()
serviceManager = ctx.ServiceManager
desktop = XSCRIPTCONTEXT.getDesktop()
doc = desktop.getCurrentComponent()
sheets = doc.getSheets() #XSpreadSheets
sheet = sheets.getByIndex(0)

def examples():


    #### operaciones básicas con celdas 

    # Obtenemos una celda y le damos valor
    xCell = sheet.getCellByPosition(0, 0)
    xCell.setValue(123)

    # Damos valor a otra celda, calculando en python con el valor de la anterior
    sheet.getCellByPosition(0,1).setValue(xCell.getValue() * 2)

    # Formula válida
    xCell = sheet.getCellByPosition(0, 2)
    xCell.setFormula("=A1+A2")

    # Formula inválida
    xCell = sheet.getCellByPosition(0, 3)
    xCell.setFormula("=1/0")
    valid = xCell.getError() == 0

    # Introducimos texto 
    xCell = sheet.getCellByPosition(1, 3)
    textCursor = xCell.createTextCursor() # xCell implementa XText
    xCell.insertString(textCursor, "La anterior celda tiene una formula erronea", False)

    # También podemos introducir texto así:
    xCell = sheet.getCellByPosition(3, 3)
    xCell.setFormula("El texto anterior es correcto")


    #### Operaciones estilos
    # Las propiedades de las celdas vienen determinadas entre otras por: 
    #    com.sun.star.table.CellProperties, com.sun.star.table.TableBorder, com.sun.star.style.ParagraphProperties, com.sun.star.style.CharacterProperties

    # Color de fondo
    xCell.setPropertyValue("CellBackColor", 0x223344); # valor en hex
     
    # Color de fuente
    xCell.setPropertyValue("CharColor", 0xFF0000);
     
    # Alineación horizontal
    #from com.sun.star.table import CellHoriJustify
    CENTER = uno.getConstantByName("com.sun.star.table.CellHoriJustify.CENTER")
    xCell.setPropertyValue("HoriJustify", CENTER)
     
    # Alineación vertical
    #from com.sun.star.table import CellVertJustify
    TOP = uno.getConstantByName("com.sun.star.table.CellVertJustify.TOP")
    xCell.setPropertyValue("VertJustify", TOP)
     

    # Estilo de fuente
    BOLD = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
    #from com.sun.star.awt import FontWeight
    xCell.setPropertyValue("CharWeight", BOLD)
    xCell.setPropertyValue("CharWeight", BOLD)
    xCell.setPropertyValue("CharWeightAsian", BOLD)
    xCell.setPropertyValue("CharWeightComplex", BOLD)
     
    # Borde 
    integerLineWidth = 100 # unidades 1/100 mm
     
    #from com.sun.star.table import BorderLine

    borderNames = ["LeftBorder", "RightBorder", "TopBorder", "BottomBorder"]
    for bname in borderNames:
        bline = uno.createUnoStruct("com.sun.star.table.BorderLine")
        #bline = BorderLine()
        bline.OuterLineWidth = integerLineWidth
        xCell.setPropertyValue(bname, bline)

    #### Operaciones con rangos
    cellAddr = uno.createUnoStruct("com.sun.star.table.CellAddress")
    cellRange = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")

    cellRange.StartColumn = 0
    cellRange.StartRow = 0
    cellRange.EndColumn = 3
    cellRange.EndRow = 3

    # Copiar
    cellAddr.Column=0
    cellAddr.Row=5
    sheet.copyRange(cellAddr, cellRange)

    cellAddr.Column=1
    cellAddr.Row=0

    # Mover
    sheet.moveRange(cellAddr, cellRange)

    # Eliminar
    cellRange.EndColumn = 0
    cellRange.EndRow = 3
    sheet.removeRange(cellRange, uno.getConstantByName("com.sun.star.sheet.CellDeleteMode.LEFT"))

    # Insertar 
    sheet.insertCells(cellRange, uno.getConstantByName("com.sun.star.sheet.CellInsertMode.RIGHT"))

    # Obtener rango por posición
    cellrng = sheet.getCellRangeByPosition(0, 0, 1, 8)
    cellrng.setPropertyValue("CellBackColor", 0x999999) # valor en hex

    # Obtener rango por nombre
    cellrng = sheet.getCellRangeByName("A1:B2")
    cellrng.setPropertyValue("CellBackColor", 0xFF9999) # valor en hex

    # Obtención de la posición del rango
    rangeAddr = cellrng.getRangeAddress()
    print(rangeAddr) # CellRangeAddress

    # Mergear celdas
    cellrng.merge(True)

    # Unmerge de celdas
    cellrng.merge(False)

    # acceso a columnas del rango
    cols = cellrng.getColumns() # XTableColumns
    ncols = cols.Count
    col = cols.getByIndex(0) 
    col.setPropertyValue("Width", 6000)

    # acceso a filas del rango
    rows = cellrng.getRows()
    rows.setPropertyValue("CellBackColor", 0xFF9999) # se aplicará a todas las celdas de las filas del rango

    # Lectura de datos en el rango
    data = cellrng.getDataArray() # Ojo! tupla de tuplas
    for row in data:
        for col in row:
            print(col)

    # Escritura de datos en el rango
    cellrng.setDataArray(((1, 3), (5, 7)))


    #### Operaciones #XSheetOperation
    res = cellrng.computeFunction(uno.getConstantByName("com.sun.star.sheet.GeneralFunction.AVERAGE"))
    print("resultado media:" + str(res))


    #### Anotaciones
    xCell = sheet.getCellByPosition(0, 0);
    aAddress = xCell.getCellAddress();

    xAnnotations = sheet.getAnnotations();
    xAnnotations.insertNew(aAddress, "This is an annotation");

    # Hacer la anotación visible
    xCell.getAnnotation().setIsVisible(True)


    #### Colecciones de celdas

    container = doc.createInstance("com.sun.star.sheet.SheetCellRanges") # XSheetCellRangeContainer
    cellRange = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    cellRange.StartColumn = 0
    cellRange.StartRow = 0
    cellRange.EndColumn = 3
    cellRange.EndRow = 3
    container.addRangeAddress(cellRange, False);

    cellRange.StartColumn = 4
    cellRange.StartRow = 4
    cellRange.EndColumn = 4
    cellRange.EndRow = 4
    container.addRangeAddress(cellRange, False);

    container.setPropertyValue("CellBackColor", 0x00FF00)

    #### Busquedas y reemplazo
    
    # buscar cadena
    xSearchDescr = sheet.createSearchDescriptor()
    xSearchDescr.SearchString = "123"
    xFound = sheet.findFirst(xSearchDescr)
    xFound.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

    # Expresión regular
    xSearchDescr.SearchString = "La ant.rior"
    xSearchDescr.SearchRegularExpression = True
    xFound = sheet.findFirst(xSearchDescr)
    xFound.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")


    #### Filtrado de columnas
    cellrng = sheet.getCellRangeByName("A1:A9")
    xFilterDesc = cellrng.createFilterDescriptor(True)
    filterField = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
    filterFields = (filterField, )
    filterField.Field = 0
    filterField.IsNumeric = True
    filterField.Operator = uno.getConstantByName("com.sun.star.sheet.FilterOperator.GREATER_EQUAL")
    filterField.NumericValue = 123

    xFilterDesc.setFilterFields(filterFields);

    xFilterDesc.setPropertyValue("ContainsHeader", True)
    cellrng.filter(xFilterDesc)


    #### Subtotales
    cellrng = sheet.getCellRangeByName("A20:A24")
    cellrng.setDataArray(((4, ),(6, ),(1, ),(-3, ),(12,)))
    xSubDesc = cellrng.createSubTotalDescriptor(True); # XSubTotalDescriptor

    aColumn = uno.createUnoStruct("com.sun.star.sheet.SubTotalColumn")
    aColumn.Column = 0;
    aColumn.Function = uno.getConstantByName("com.sun.star.sheet.GeneralFunction.SUM")
    aColumns = (aColumn, )

    xSubDesc.addNew(aColumns, 1); # columna en la que se ponen textos de totales (respeta los textos que haya)
    cellrng.applySubTotals(xSubDesc, True);

