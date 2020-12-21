printerURL = "\\printserver\printshare"

Set objNetwork = WScript.CreateObject("WScript.Network")
objNetwork.RemovePrinterConnection printerURL