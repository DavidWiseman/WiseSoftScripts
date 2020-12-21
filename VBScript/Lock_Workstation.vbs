Set objShell = CreateObject("WScript.Shell")

objShell.Run "RunDll32.exe user32.dll,LockWorkStation"
