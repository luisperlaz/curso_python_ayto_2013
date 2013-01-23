# -*- coding: utf-8 -*-
import uno
import sys

def mostrar_version():  
    # El documento desde donde se llama esta macro
    oDoc = XSCRIPTCONTEXT.getDocument()
    # El manejador de servicios
    oSM = uno.getComponentContext().getServiceManager()
    # Creamos una instancia del servicio Toolkit
    oToolkit = oSM.createInstance( "com.sun.star.awt.Toolkit" )
    # Creamos una estructura Rectangulo
    rec = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    # El tipo de cuadro de mensaje, solo informativo
    sTipo = "infobox"
    # El Título del cuadro de mensaje
    sTitulo = "Informativo"
    # El mensaje
    sMensaje = "Versión de python: " + str(sys.version)
    # Mostraremos solo el botón Aceptar
    botones = 1
    # Referencia a la ventana contenedora
    oParentWin = oDoc.getCurrentController().getFrame().getContainerWindow()
    # Creamos el cuadro de mensaje con los parámetros necesarios
    oMsgBox = oToolkit.createMessageBox( oParentWin, rec, sTipo, botones, sTitulo, sMensaje )
    # Mostramos el cuadro de mensaje
    oMsgBox.execute()
    return None

