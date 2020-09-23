Attribute VB_Name = "modRooms"
' This is for handling of the room data and other stuff

Option Compare Text
' The latest user to enter NChat
Public NewUser As String
' The room's welcome message
Public WelcomeMsg As String
' The name of the room you are in
Public RoomName As String
' Your room description
Public Description As String
' Have you created your own room?
Public CreatedRoom As Boolean
' How long your room has been running
Public RoomTime As Long
' Your room "ID" (Port)
Public RoomID As Long
' Umm.... I think this is the same as above. Don't
' ask why. It just is...
Public Room As Long
Public RoomHost As String
' The password for any newly created rooms
Public Password As String


Public Sub NewRoom(ByVal RoomPort As Long, Optional RoomN As String, Optional Silent As Boolean)
' New Room is actually a port change.
' This may interfere with other applications, but
' That aint my fault :P
    On Error Resume Next
    DoEvents
    
    ' If your on a server, then don't allow room changes, because this
    ' could be a security risk
    If IsInternet = True Then
        MsgBox "Cannot change rooms! Your on a server! Disconnect from the server (From the main menu) and try again", vbCritical, "Unable to change rooms!"
        Exit Sub
    End If
    
    ' No use in disconnecting yourself if no
    ' username has been set
    If UserName > "" Then
        Broadcast "disø" & UserName
        DoEvents
        DoEvents
    End If
    ' Resets your "Created room" status
    CreatedRoom = False
    Password = ""

    DoEvents

    DoEvents
    Room = RoomPort
    ' Clear the user list and welcome message
    frmMain.List1.ListItems.Clear
    WelcomeMsg = ""
    Description = ""
    ' Rests the connection, but this time with
    ' the Room ID as the port

    frmMain.sckUDP.Close

    frmMain.sckUDP.RemoteHost = Address
    frmMain.sckUDP.RemotePort = RoomPort
    frmMain.sckUDP.LocalPort = frmMain.sckUDP.RemotePort



    If UserName > "" Then
        Broadcast ("conø" & MyIcon & "ø" & sckUDP.LocalIP)
        DoEvents
        DoEvents
        Text "+username+ has entered the conversation..." & vbCrLf, "con", True, , , , "Center"
    End If

    DoEvents
    If RoomN = "" And Silent = False Then
        Text "  >> Connected to Room #" & RoomPort & ". Enjoy!" & vbCrLf, "con", True
    ElseIf RoomN > "" And Silent = False Then
        Text "  >> Connected to " & RoomN & " (Room #" & RoomPort & ") Enjoy!" & vbCrLf, "con", True
        RoomName = RoomN
    End If

    ' Silent doesn't tell you when you have changed rooms
    DoEvents
    ' Add your name, and let the rest of the room know
    ' you have arrived
    frmMain.List1.ListItems.Add 1, frmMain.sckUDP.LocalIP, UserName, , frmMain.ImageList1.ListImages.Item(FindIcon(MyIcon)).Key
    DoEvents
    DoEvents

End Sub

