import uno
import unohelper
import time
from com.sun.star.task import XJobExecutor
 
class Neojob( unohelper.Base, XJobExecutor ):
    def __init__( self, ctx ):
        self.ctx = ctx
 
    def trigger( self, args ):
        print("executing.... %s" % str(args))
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
                found = doc.findNext( found.End, search)
 
        except:
            pass

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
        Neojob,
        "es.neodoo.libreoffice.example.Neojob",
        ("com.sun.star.task.Job",),)
