Attribute VB_Name = "modZUser"
Public Type RemoteUserDetails
    RUserName As String
    RIPAddress As String
    RSHAHash As String
End Type

Public Function FindUser(UserName As String) As Integer
' Find a username and return it's index in the frmMain.list1
    For n = 1 To frmMain.List1.ListItems.Count
        If frmMain.List1.ListItems(n).Text = UserName Then
            FindUser = n
            Exit Function
        End If
    Next n
End Function

Public Sub AddUser(name As String, Icon As String, IP As String)
    Dim tempint As Integer
    On Error Resume Next
    tempint = FindIcon(Icon)
    ' Can't find an icon? Then use the default one
    If tempint = 0 Then tempint = 1
    frmMain.List1.ListItems.Add , IP, name, , frmMain.ImageList1.ListImages.Item(tempint).Key
    '    frmMain.List1.ListItems.Item(2).SubItems(1) = 0
End Sub


Public Sub RemoveUser(name As String)
' Find the user using the FindUser function (above), which returns an integer,
' which can be used
    If FindUser(name) > 0 Then frmMain.List1.ListItems.Remove (FindUser(name))
End Sub

Public Function FindIcon(name As String) As Integer
    For x = 1 To frmMain.ImageList1.ListImages.Count
        If frmMain.ImageList1.ListImages.Item(x).Key = name Then
            FindIcon = x
            Exit Function
        End If
    Next
    x = -1
End Function
