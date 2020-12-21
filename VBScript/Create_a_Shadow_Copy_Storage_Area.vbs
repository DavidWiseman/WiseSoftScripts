Const VOLUME = "C:\"
Const DIFFERENTIAL_VOLUME = "E:\"
Const MAXIMUM_SPACE = 130023424
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objShadowStorage = objWMIService.Get("Win32_ShadowStorage")
errResult = objShadowStorage.Create(VOLUME, DIFFERENTIAL_VOLUME, MAXIMUM_SPACE)
