Set objNet = CreateObject("WScript.Network" ) 

' Set the user you want to make local administrator here 
strUser = "<username here>"

strNetBIOSDomain = objNet.UserDomain 
strComputer = objNet.ComputerName 

Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group" ) 
Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" & strUser & ",user" ) 

' ignore error if user is already a member of the group
On Error Resume Next 
objGroup.Add(objUser.ADsPath) 
On Error Goto 0 
