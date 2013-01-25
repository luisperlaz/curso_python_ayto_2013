# -.- coding: utf-8 -.-

import uno
import unohelper
import os

from com.sun.star.awt import XDialogEventHandler

ctx = XSCRIPTCONTEXT.getComponentContext()
serviceManager = ctx.ServiceManager
desktop = XSCRIPTCONTEXT.getDesktop()
doc = desktop.getCurrentComponent()



class Handler(unohelper.Base, XDialogEventHandler):
    def callHandlerMethod(self, xDialog, event, methodName):
        print("invokedd")
        xTextField1Control = xDialog.getControl("TextField1")
        print("texto obtenido")
        xControlModel1 = xTextField1Control.getModel()
        print("modelo obtenido")
        aText = xControlModel1.getPropertyValue("Text")
        print("aText: " + aText)
        
        xDialog.endExecute()
        if doc:
            text = doc.Text
            cursor = text.createTextCursor()
            text.insertString( cursor, aText, 0 )
        

    def getSupportedMethodNames(self):
        print("getting names")
        return ("handleButton",)

    def handleButton(self, xDialog, event):
        print("eoo")

        

def namedialog():

    args = (doc,)
    dialogprov = serviceManager.createInstanceWithArgumentsAndContext("com.sun.star.awt.DialogProvider2", args, ctx)
    #dialog = dialogprov.createDialogWithHandler("file:///tmp/NameDialog.xdl", Handler()) # También se puede referenciar el archivo del diálogo
    dialog = dialogprov.createDialogWithHandler("vnd.sun.star.script:NeoLibrary.NameDialog?location=application", Handler())
    dialog.execute()
    

