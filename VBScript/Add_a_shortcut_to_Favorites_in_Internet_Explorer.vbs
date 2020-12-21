Const ADMINISTRATIVE_TOOLS = 6

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS) 
Set objFolderItem = objFolder.Self     

Set objShell = WScript.CreateObject("WScript.Shell")
strDesktopFld = objFolderItem.Path
Set objURLShortcut = objShell.CreateShortcut(strDesktopFld & "\WiseSoft.url")
objURLShortcut.TargetPath = "http://www.wisesoft.co.uk"
objURLShortcut.Save
