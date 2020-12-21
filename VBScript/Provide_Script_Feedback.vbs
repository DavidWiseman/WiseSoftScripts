'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Edit constants to change the HTML display
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
const company = "WiseSoft"
const companyURL = "http://www.wisesoft.co.uk"
const creator = "David Wiseman"
const scriptName = "Script to automate everything on my network"
const website = "www.wisesoft.co.uk"
const createdDate = "10/06/2005"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Open IE to display the script progress
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
set objExplorer = createObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"
objExplorer.Toolbar =0
objExplorer.Statusbar = 0
objExplorer.Width = 600
objExplorer.Height = 300
objExplorer.Left = 0
objExplorer.Top = 0

do while (objExplorer.Busy)
	wscript.sleep 200
loop

objExplorer.Visible = 1

BodyText = "<h2><a href = '" & companyURL & "' target='_blank'>" & _
	   company & "</a></h2>" & "<b>Script Function:</b> " & _
	   scriptName & "<BR><b>Created By</b> " & creator & _
	   "</font><BR><b>Created:</b> " & createdDate & _
	   "<BR><HR><BR><b><font color=#9932cd><i>Script Progress:</i></font></b> "
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' This is the body of your script.  Insert your script code here.
' Add the line of code that changes the HTML to key points in your
' script.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

objExplorer.Document.Body.InnerHTML = BodyText & "Starting up..."

wscript.sleep 5000

objExplorer.Document.Body.InnerHTML = BodyText & "Doing my magic..."

wscript.sleep 7000

objExplorer.Document.Body.InnerHTML = BodyText & "Cleaning up..."

wscript.sleep 1500

objExplorer.Document.Body.InnerHTML =  BodyText & "Finished..."


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' The following code will close the IE window after the script
' has finished - a wait of 10 seconds has been added.  On error
' resume next is used as user might have closed IE.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

wscript.sleep 10000
on error resume next
objExplorer.Quit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~