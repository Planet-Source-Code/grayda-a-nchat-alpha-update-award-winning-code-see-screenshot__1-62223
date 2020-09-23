Attribute VB_Name = "modWinsock"
Public Sub SendData(Data As String, Winsock_Number As Integer, Optional SendToAll As Boolean = False, Optional NumberOfWinsocks As Integer)
If SendToAll = True Then
    For i = 1 To NumberOfWinsocks
        frmMain.sckServer(i).SendData Encode(Data, frmMain.txtServerPassword.Text)
    Next i
Else
    frmMain.sckServer(Winsock_Number).SendData Encode(Data, frmMain.txtServerPassword.Text)
End If
    

End Sub
