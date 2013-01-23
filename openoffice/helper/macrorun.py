#!/usr/bin/env python

import uno
import os
import subprocess
import argparse
from com.sun.star.connection import NoConnectException


def getServiceManager(host="localhost", port="2002"):
    ctx = get_context(host, port)
    return ctx.ServiceManager

def get_resolver(local_ctx):
    return local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)

def get_context(host="localhost", port="2002"):
    local_ctx = uno.getComponentContext()
    resolver = get_resolver(local_ctx)
    connurl ="uno:socket,host={},port={};urp;StarOffice.ComponentContext"
    return resolver.resolve(connurl.format(host, port))
    

def get_desktop(serviceManager, context):
    return serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

def loadComponentFromURL(desktop, cUrl, tProperties=() ):
    """Open or Create a document from it's URL."""
    oDocument = desktop.loadComponentFromURL(cUrl, "_blank", 0, tProperties)
    return oDocument

def createNewWriterDoc(desktop):
    loadComponentFromURL("private:factory/swriter")

def createNewCalcDoc(desktop):
    loadComponentFromURL("private:factory/scalc")

def startOffice(host="localhost", port="2002"):
    NoConnectionException = uno.getClass("com.sun.star.connection.NoConnectException")
    ooffice = 'soffice --norestore "-accept=socket,host={},port={};urp;"'.format(host, port)
    if os.fork():
        return

    retcode = subprocess.call(ooffice, shell=True)
    if retcode != 0:
        print("OOo returned {}".format(retcode))#, file=sys.stderr)


class XScriptContext(object):
    def __init__(self, document, desktop, ctx):
        self.document = document
        self.desktop = desktop
        self.ctx = ctx

    def getDocument(self):
        return self.document

    def getDesktop(self):
        return self.desktop

    def getComponentContext(self):
        return self.ctx

def _create_scriptcontext(host, port):
    ctx = get_context(args.host, args.port)
    desktop = get_desktop(ctx.ServiceManager, ctx)
    doc = desktop.getCurrentComponent()

    oToolkit = ctx.ServiceManager.createInstance( "com.sun.star.awt.Toolkit" )

    return XScriptContext(doc, desktop, ctx)


if __name__ == "__main__":


    parser = argparse.ArgumentParser(description='Execute script as if it were a macro.')
    parser.add_argument('scriptfile', action='store') #, type=argparse.FileType())
    parser.add_argument('macrofunc', action='store') 
    parser.add_argument('-H', '--host', action='store', default="localhost", help="The host to connect to. Default: localhost")
    parser.add_argument('-p', '--port', action='store', default="2002", help="The port to connect to. Default: 2002")
    
    args = parser.parse_args()

    try: 
        XSCRIPTCONTEXT = _create_scriptcontext(args.host, args.port)
        execfile(args.scriptfile)
        eval(args.macrofunc+"()")
    except NoConnectException, e:
        cmd = 'soffice --norestore "-accept=socket,host={},port={};urp;"'.format(args.host, args.port)
        print("Unable to connect to libreoffice. Is it running?")
        print("You can start it like this: " + cmd)
        exit(-1)



