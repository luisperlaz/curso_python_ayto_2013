import uno
 
def HolaMundoCalc():
	# Accedemos al modelo del documento actual 
	model = XSCRIPTCONTEXT.getDocument()
	# Accedemos a la primer hoja del documento
	hoja = model.getSheets().getByIndex(0)
	# Accedemos a la celda A1 de la hoja
	celda = hoja.getCellRangeByName("A1")
	# Escribimos en la celda
	celda.setString("Hola Mundo en Python")
	return None

