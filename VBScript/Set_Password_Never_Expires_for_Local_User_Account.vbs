Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000

strComputer = "."
strUser = "User01"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
intFlags = objUser.Get("UserFlags")
intFlags = intFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", intFlags 
objUser.SetInfo