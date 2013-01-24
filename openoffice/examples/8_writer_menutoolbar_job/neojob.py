import uno
import unohelper
import time
from com.sun.star.task import XJob
 
class Neojob( unohelper.Base, XJob ):
    def __init__( self, ctx ):
        self.ctx = ctx
 
    def execute(self, args):
        print("executing....")
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx )
 
        doc = desktop.getCurrentComponent()

 
        try:
            search = doc.createSearchDescriptor()
            search.SearchRegularExpression = True
            search.SearchString = "foo"
 
            found = doc.findFirst( search )
            while found:
                found.String = found.String.replace("foo", u"bar" )
                time.sleep(3)
                found = doc.findNext( found.End, search)
 
        except:
            pass
        return -1

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
        Neojob,
        "es.neodoo.libreoffice.example.Neojob",
        ("com.sun.star.task.Job",),)
