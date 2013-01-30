# -.- coding: utf-8 -.-

from urllib2 import urlopen
from json import load
 

def _get_estaciones():
	urlt = 'http://www.zaragoza.es/buscador/select?q=category:Bizi&wt=json'
	res = load(urlopen(urlt))
	estaciones = res['response']['docs']
	return estaciones

def _get_estaciones_tuples():
	tuples = []
	estaciones = _get_estaciones()
	for estacion in estaciones:
		localizacion = estacion['title']
		bicis = estacion['bicisdisponibles_i']
		anclajes = estacion['anclajesdisponibles_i']
		estado = estacion['estado_t']
		tuples.append((localizacion, bicis, anclajes, estado))

	return tuple(tuples)

def create_bizi_calc():

	estaciones = _get_estaciones()

	ctx = XSCRIPTCONTEXT.getComponentContext()
	serviceManager = ctx.ServiceManager
	desktop = XSCRIPTCONTEXT.getDesktop()
	doc = desktop.getCurrentComponent()
	sheets = doc.getSheets() #XSpreadSheets
	sheet = sheets.getByIndex(0)

	# popular datos
	cellrng = sheet.getCellRangeByPosition(0, 0, 3, len(estaciones)-1)
	data = _get_estaciones_tuples()
	cellrng.setDataArray(data)

	# insertar fila con títulos
	sheet.getRows().insertByIndex(0, 1)
	titles = sheet.getRows().getByIndex(0).getCellRangeByPosition(0,0,3,0)
	titles.setDataArray((("Localizacion", "bicis disponibles", "anclajes disponibles", "estado"),))

	# dar ancho a columna de titulo
	col1 = sheet.getColumns().getByIndex(0)
	col1.setPropertyValue('Width', 8000)

	# dar estilo a titulos
	titles.setPropertyValue('CellBackColor', 0x888888)
	
	borderNames = ["LeftBorder", "RightBorder", "TopBorder", "BottomBorder"]
    	for bname in borderNames:
		bline = uno.createUnoStruct("com.sun.star.table.BorderLine")
		bline.OuterLineWidth = 50
		titles.setPropertyValue(bname, bline)

	# añadir celda de suma de bicis disponibles
	nestaciones = len(estaciones)
	cellsum = sheet.getCellByPosition(1, nestaciones + 1)
	cellsum.setFormula("=SUM(B2:B{})".format(nestaciones + 1))


	valuesrng = sheet.getCellRangeByPosition(1, 1, 2, nestaciones)
	xEntries = valuesrng.getPropertyValue("ConditionalFormat");

	cond1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
	cond2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
	cond3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")

	cond1.Name = "Operator";
	cond1.Value = uno.getConstantByName("com.sun.star.sheet.ConditionOperator.EQUAL")
	cond2.Name = "Formula"
	cond2.Value = "0"
	cond3.Name = "StyleName"
	cond3.Value = "incorrecto"

	xEntries.addNew((cond1, cond2, cond3))
	valuesrng.setPropertyValue("ConditionalFormat", xEntries);

