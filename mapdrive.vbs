Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("Wscript.Network")

Call MapDrive("h:", "\\computer\share", "False")

Function MapDrive(strDrive, strShare, bReplace)
' Function to map network share to a drive letter.
' If the drive letter specified is already in use, the function
' attempts to remove the network connection.
' objFSO is the File System Object, with global scope.
' objNetwork is the Network object, with global scope.
' Returns True if drive mapped, False otherwise.

  Dim objDrive

  On Error Resume Next
  Err.Clear
  If objFSO.DriveExists(strDrive) Then
    Set objDrive = objFSO.GetDrive(strDrive)
    If Err.Number <> 0 Then
      Err.Clear
      MapDrive = False
      showstat("Unable to Map "& strDrive &" to "& strShare) 
      Exit Function
    End If
    If objDrive.ShareName = strShare Then
      'Ignore Correctly Mapped Network Drives
      'Wscript.Echo strShare & vbTab & objDrive.ShareName      
      MapDrive = True
      Exit Function
    ElseIf objDrive.ShareName <> strShare AND bReplace = "True" Then
      objNetwork.RemoveNetworkDrive strDrive, True, True
      MapDrive = True
      showstat("Override Mapping "& strDrive &" to "& strShare) 
    Else
      MapDrive = False
      showstat("Unable to Map "& strDrive &" to "& strShare) 
      Exit Function
    End If
    Set objDrive = Nothing
  End If
  objNetwork.MapNetworkDrive strDrive, strShare
  If Err.Number = 0 Then
    MapDrive = True
  Else
    Err.Clear
    MapDrive = False
    showstat("Unable to Map "& strDrive &" to "& strShare) 
  End If
  On Error GoTo 0
End Function
