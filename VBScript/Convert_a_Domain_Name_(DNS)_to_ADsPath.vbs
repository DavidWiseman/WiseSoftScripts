strDomainName = "scripting.wisesoft.co.uk"
arrDomLevels = Split(strDomainName, ".")
strADsPath = "dc=" & Join(arrDomLevels, ",dc=")
WScript.Echo strADsPath