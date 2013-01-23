import uno
 
""" Here is the sequence of things the lines do:
1.  Get the uno component context from the PyUNO runtime
2.  Create the UnoUrlResolver
3.  Get the central desktop object
4.  Declare the ServiceManager
5.  Get the central desktop object
6.  Access the current writer document
7.  Access the document's text property
8.  Create a cursor
9.  Insert the text into the document """
 
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
model = desktop.getCurrentComponent()
text = model.Text
cursor = text.createTextCursor()
text.insertString( cursor, "Hello World", 0 )
 
""" Do a nasty thing before exiting the python process. In case the
 last call is a one-way call (e.g. see idl-spec of insertString),
 it must be forced out of the remote-bridge caches before python
 exits the process. Otherwise, the one-way call may or may not reach
 the target object.
 I do this here by calling a cheap synchronous call (getPropertyValue)."""
ctx.ServiceManager
