Const VOLUME = "C:\"
Const CONTEXT = "ClientAccessible"
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objShadowStorage = objWMIService.Get("Win32_ShadowCopy")
errResult = objShadowStorage.Create(VOLUME, CONTEXT, strShadowID)
