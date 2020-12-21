$ie = new-object -comobject "InternetExplorer.Application"

$ie.visible = $true

$ie.navigate("www.wisesoft.co.uk")
