Option Explicit

Dim objWshNetwork
Set objWshNetwork = WScript.CreateObject("WScript.Network")

Dim objFSO
Set objFSO WScript.CreateObject("Scripting.FileSystemObject")

Dim strShareFolders
strShareFolders = Array("\\10.10.10.10\c$")
'strShareFolders = Array("\\10.10.10.10\work")

Dim intChr, intDrive, strDrive
intDrive = 0

For intChr = Asc("C") To Asc("Z")
	strDrive = Chr(intChr) & ":"
	If Not objFSO.DriveExists(strDrive) Then
		If intDrive <= UBound(strShareFolders) Then
			objWshNetwork.MapNetworkDrive strDrive, strShareFolders(intDrive), False, "userid", "password"
		End If
	End If
Next

Set objWshNetwork = Nothing
