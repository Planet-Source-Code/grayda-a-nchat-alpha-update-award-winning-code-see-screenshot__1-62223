Attribute VB_Name = "modWSock"
Option Compare Text

' Is NChat running on a LAN (ie. Connectionless UDP)
' or is it running over the internet (ie. TCP/IP Connection)
Public IsInternet As Boolean

' Data recieved through sckUDP (in frmMain
Public sData As String
' Results of the SplitVB5 function
Public Result() As String
' The sckUDP and sckRooms addresses that are to be used
' They are string because of the decimal points inbetween
Public Address As String

Public Sub Broadcast(CData As String)
    On Error Resume Next
    ' Sends data to EVERYONE in the room

    ' These are text shortcuts. When you type these,
    ' they are replaced with correct values. this makes
    ' chatting easier, because Welcome Messages, Notch
    ' and other stuff can be more dynamic

    ' No Data? Don't send it
    If CData = "" Then Exit Sub

    CData = Replace(CData, "+username+", UserName)
    CData = Replace(CData, "+ip+", frmMain.sckUDP.LocalIP)
    CData = Replace(CData, "+room+", RoomName & " (" & Room & ")")
    CData = Replace(CData, "+ncredits+", NCredits)
    CData = Replace(CData, "+newuser+", NewUser)
    CData = Replace(CData, "+ver+", Ver)
    CData = Replace(CData, "+ntime+", NChatTime)
    CData = Replace(CData, "+time+", Format(Time, "HH:mm"))
    ' +d+ is our delimiter for sent data. This is simpler
    ' than typing Alt+0248
    CData = Replace(CData, "+d+", "ø")
    ' Who our last message is from
    CData = Replace(CData, "+lastfrom+", LastFrom)
    CData = Replace(CData, "+roomhost+", RoomHost)
    Randomize
    CData = Replace(CData, "+someguy+", frmMain.List1.ListItems.Item(Int(Rnd * frmMain.List1.ListItems.Count) + 1).Text)

    FindPos = InStr(1, CData, "+result")
    If FindPos > 0 Then CData = Replace(CData, "+result" & Mid(CData, FindPos + Len("+result"), 1) & "+", Result(Mid(CData, FindPos + Len("+result"), 1)))

    ' Need to insert Smiley Locator Code HEre

    If IsInternet = False Then
        frmMain.sckUDP.Close
        ' 'Resets' the connection
        frmMain.sckUDP.LocalPort = frmMain.sckUDP.RemotePort
        frmMain.sckUDP.RemoteHost = Address
        frmMain.sckUDP.RemotePort = Room

        frmMain.sckUDP.Connect
    End If

    ' Finally sends the data
    ' N©H@-|- is our data 'header', and a SHA Hash code is included for verification...
    
    ' If the incoming data is missing this, then
    ' ignore it
    Dim ToSend As String
    Dim HashTemp As String
       
    ToSend = "N©H@-|-ø" & CData
    ToSend = ToSend & "ø" & frmMain.sckUDP.LocalIP & "ø" & UserName
        
    HashTemp = SHAHash(ToSend)
    frmMain.sckUDP.SendData ToSend & "ø" & HashTemp
    If Left(CData, 4) = "Move" Or Left(CData, 3) = "usr" Or Left(CData, 5) = "Clear" Or Left(CData, 6) = "Colour" Or Left(CData, 4) = "Size" Then Exit Sub
    If Trim(CData) > "" Then
     frmLog.ListView1.ListItems.Add , , CData
     frmLog.ListView1.ListItems(frmLog.ListView1.ListItems.Count).SubItems(1) = frmMain.sckUDP.LocalIP
     frmLog.ListView1.ListItems(frmLog.ListView1.ListItems.Count).SubItems(2) = HashTemp
     frmLog.ListView1.ListItems(frmLog.ListView1.ListItems.Count).SubItems(3) = UserName
     
    '    frmLog.List2.AddItem CData
    Else
        frmLog.ListView1.ListItems.Add , , "--Blank Data--"
    End If


End Sub

Public Sub RoomBroadcast(BText As String)
' This is to do with the list of rooms
' See the sckRooms_Dataarival sub for more info

' 'Resets' the connection
    On Error Resume Next
    If IsInternet = False Then
        frmMain.sckRooms.Close
        frmMain.sckRooms.RemoteHost = Address
        frmMain.sckRooms.LocalPort = 127
        frmMain.sckRooms.RemotePort = 127
    End If
    'frmMain.sckRooms.Connect
    ' Finally sends the data
    frmMain.sckRooms.SendData BText

End Sub
