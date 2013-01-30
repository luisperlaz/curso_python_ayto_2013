import uno
import unohelper
from es.neodoo.libreoffice.examples.IVA import XIVA

# Impl. de IVA

class IvaImpl(unohelper.Base, XIVA):
	def __init__(self, ctx):
		self.ctx = ctx

	def IVA(self, precio):
		return precio * 1.21
 
def createInstance(ctx):
	return IvaImpl(ctx)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation( \
	createInstance,"es.neodoo.libreoffice.examples.IvaImpl",
		("com.sun.star.sheet.AddIn",),)
