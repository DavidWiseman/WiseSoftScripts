On Error Resume Next

DIM objNetwork,colDrives,i

Set objNetwork = CreateObject("Wscript.Network")

Set colDrives = objNetwork.EnumNetworkDrives

For i = 0 to colDrives.Count-1 Step 2
	' Force Removal of network drive and remove from user profile 
	' objNetwork.RemoveNetworkDrive strName, [bForce], [bUpdateProfile]
	objNetwork.RemoveNetworkDrive colDrives.Item(i),TRUE,TRUE
Next

