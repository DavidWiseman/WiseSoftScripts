PrintServer = "\\printserver\printshare"

set objNetwork = createobject("Wscript.Network")

objNetwork.AddWindowsPrinterConnection(PrintServer)

'**** Make the printer default
objNetwork.SetDefaultPrinter PrintServer