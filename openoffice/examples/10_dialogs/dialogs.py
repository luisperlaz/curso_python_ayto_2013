import uno

from com.sun.star.awt.MessageBoxButtons import (BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, 
BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY, DEFAULT_BUTTON_OK, 
DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE)

from com.sun.star.awt import Rectangle

def CuadroMensaje3():  
    # El documento desde donde se llama esta macro
    oDoc = XSCRIPTCONTEXT.getDocument()
    # El manejador de servicios
    oSM = uno.getComponentContext().getServiceManager()
    # Creamos una instancia del servicio Toolkit
    oToolkit = oSM.createInstance( "com.sun.star.awt.Toolkit" )
    # Creamos una estructura Rectangulo
    rec = Rectangle()
    # El tipo de cuadro de mensaje, solo informativo
    sTipo = "warningbox"
    # El Título del cuadro de mensaje
    sTitulo = "Precaución"
    # El mensaje
    sMensaje = "Hay un problema ¿Que deseas hacer?"
    # Mostraremos solo el botón Aceptar
    botones = BUTTONS_ABORT_IGNORE_RETRY
    # Referencia a la ventana contenedora
    oParentWin = oDoc.getCurrentController().getFrame().getContainerWindow()
    # Creamos el cuadro de mensaje con los parámetros necesarios
    oMsgBox = oToolkit.createMessageBox( oParentWin, rec, sTipo, botones, sTitulo, sMensaje )
    # Mostramos el cuadro de mensaje
    res = oMsgBox.execute()
    # Determinamos que botón presiono el usuario
    if res == 0 :
        sMensaje = "El usuario presiono Cancelar"
    elif res == 4 :
        sMensaje = "El usuario presiono Repetir"
    elif res == 5 :
        sMensaje = "El usuario presiono Ignorar"

    sMensaje = "%s code %s" % (sMensaje, DEFAULT_BUTTON_CANCEL)
    # Mostramos el mensaje
    oMsgBox = oToolkit.createMessageBox( oParentWin, rec, "infobox", botones, "Respuesta", sMensaje )
    oMsgBox.execute()
    return None

