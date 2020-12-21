Set objDlg = CreateObject("UserAccounts.CommonDialog") 
objDlg.Filter = "All Files|*.*"
blnReturn = objDlg.ShowOpen 

if blnReturn then 
	WScript.Echo objDlg.FileName 
else
	wscript.echo "Open File Dialog Cancelled"
end if