Attribute VB_Name = "modPasrse"
' OK, here's the new layout of the data. It serves 2 purposes:
' 1) To stop backwards compatibility (Sorry, but a lot of the new stuff simply
'    won't be compatible with the old stuff.
' 2) To ensure simplicity of data parsing etc.

' Some Dims. This is for the SplitUpData thingy:
Dim DataCode As String, UID As String, Params() As String, Message As String, TempArray() As String



Public Sub SplitUpData(Data As String)
' OK this lets us split up our new data, into the 4 main parts:
' Data Code (or prefix)
' Unique User ID
'

Dim Temp As Integer, Temp1 As Integer, Temp2 As Integer, Temp3 As Integer

' Part 1, from the start of the string to the ø
Temp = InStr(1, Data, "ø") - 1
DataCode = Left(Data, Temp)
'MsgBox DataCode
Temp1 = InStr(Temp, Data, "!")
UID = Mid(Data, Temp + 2, Temp1 - Temp - 2)
'MsgBox UID

Temp2 = InStr(Temp1, Data, "@")
Message = Right(Data, Len(Data) - InStrRevVB5(Data, "@"))
'Params = SplitVB5(Mid(Data, Temp1, Len(Data) - InStrRevVB5(Data, "@")), ",")
Params = SplitVB5(Mid(Data, Temp1 + 1, Len(Data) - Temp2 + Len(Message) - 1), ",")

End Sub

Public Sub ParseData(Data As String)

      Dim SHA_Hash As String
'10    On Error GoTo ErrorH


          ' When we send the data, NChat slaps a 40 character-long SHA Hash on the end
          ' for the other side to verify. To strip the hash off the end, get everything
          ' on the left hand side, except for the last 40 characters
20    SHA_Hash = Right(sData, 40)
        
        
        
          ' Once the first message has arrived, set the text again
          ' so no errors appear
30    frmMain.SB1.Panels(1).Text = "NChat - Online!!"
40    frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture
50    tempdata = sData

          ' Excuse the language. If swearing is off, then replace
          ' the language with a cool icon that says "Censored"
60    If Swearing = False Then
70      sData = Replace(sData, "shit", ":censored")
80      sData = Replace(sData, "fuck", ":censored")
90      sData = Replace(sData, "cunt", ":censored")
100     sData = Replace(sData, "bitch", ":censored")
110   End If

120   OldData = sData

DoEvents
DoEvents

130   Result = SplitVB5(sData, "ø")
140   atemp = UBound(Result)
      'MsgBox sData
      Dim ToHash As String

150   ToHash = Left(sData, Len(sData) - Len(Result(atemp)) - Len(Result(atemp - 1)) - Len(Result(atemp - 2)) - 4)
      'SHA_Hash = SHAHash(ToHash)
      'Text SHA_Hash & vbCrLf & ToHash & vbCrLf
      'Text Result(atemp) & vbCrLf
160   If Result(atemp) <> SHA_Hash Then
      'Text "Error" & vbCrLf, "ThatsBad"
170   MsgBox "This data has been modified. Expected " & Result(atemp) & " but got " & SHA_Hash
180   End If

          ' No Blank packets of data thanks! btw, N©H@-|- is to ensure that
          ' other programs don't accidentally interfere with NChat on UDP mode
190   If Trim(sData) = "" Or Result(0) <> "N©H@-|-" Then Exit Sub

200   sData = Replace(sData, "+newline+", vbCrLf)

          ' +D+ is short for delimiter. Less keystrokes than pressin
          ' alt+0248
210   sData = Replace(sData, "+d+", "ø")
220   RUser.RIPAddress = frmMain.sckUDP.RemoteHostIP
230   RUser.RUserName = Result(atemp - 1)
240   RUser.RSHAHash = SHA_Hash
250   DoEvents
          ' Before NChat v4.0 (about March 2004), I didn't know
          ' what the Split command did! I have been programming
          ' for about 6-7 years, and only learnt in 2004
          ' what it does! Before that, you could only
          ' have commands that consisted of two parts.
          ' The first part MUST be 3 letters long, and the
          ' second part could be any amount of letters long.
          ' Except in exceptional circumstances, you couldn't
          ' have a third or fourth part :|

260   Select Case Result(1)    ' Result(1) can be any length now

          Case "msg"
        ' Result(2) is the username and Result(3) is the message
        ' Rather than sending || Username || Hello as the message,
        ' The ||'s are added on arrival

        ' Remember: StartMSG = "||" until changed

        ' OK this section has been updated. If you
        ' purchase the items from the store, then you can
        ' have a bold, underlined, or even a different
        ' coloured text for your username

        ' This ensures Notch DOESN'T Talk until
        ' noone has talked for a certain amount of time
270     If frmAutoBotOptions.Timer1.Enabled = True Then
280         frmAutoBotOptions.Timer1.Enabled = Not frmAutoBotOptions.Timer1.Enabled
290         frmAutoBotOptions.Timer1.Enabled = Not frmAutoBotOptions.Timer1.Enabled
300     End If

310     If FileObj.FileExists(Left(Profile, Len(Profile) - Len("index.htm")) & "messages.htm") = True Then
320         Open Left(Profile, Len(Profile) - Len("index.htm")) & "messages.htm" For Input As #1
330         Text Input(LOF(1), 1)
340         Close #1
350     End If
360     If FileObj.FileExists(Left(Profile, Len(Profile) - Len("index.htm")) & "messages.htm") = False Then
370         Text StartMSG & " ", "Msg", False, False, False, 2
380         Text RUser.RUserName & " ", Result(3), CBool(Result(4)), False, CBool(Result(5)), 2
390         Text EndMSG & " ", "Msg", False, False, False, 2
400     End If
410   Text Result(2) & vbCrLf
420     LastFrom = RUser.RUserName

        ' Log Logs the data into the frmLog textbox
430     Log "Message: " & RUser.RUserName & " - " & Result(3) & vbCrLf, vbBlue
        ' Changes the caption to say the new message
440     Status CaptionPrefix & Result(2) & " - " & Result(3) & EndWindow

450   Case "act"    ' action
460     If frmAutoBotOptions.Timer1.Enabled = True Then
470         frmAutoBotOptions.Timer1.Enabled = False
480         frmAutoBotOptions.Timer1.Enabled = True
490     End If

500     Text RUser.RUserName & Result(2) & vbCrLf, "act", True, , , 2
510     Status CaptionPrefix & "A: " & RUser.RUserName & Result(2) & EndWindow
520     Log "Action: " & RUser.RUserName & Result(2) & vbCrLf, 33023

530   Case "svr"    ' Server (Purple) message
540     Text Result(2) & vbCrLf, svr, True, , , 2
550     Status CaptionPrefix & "S: " & Result(2) & EndWindow
560     Log "Server: " & Result(2) & vbCrLf, vbMagenta

570   Case "add"    ' Add administrator
580     Log "Add Admin: " & Result(2) & " added by: " & RUser.RUserName & vbCrLf, 32768

590     If Result(2) = UserName Then
600         frmMain.mnuDev.Visible = True
610         Text RUser.RUserName & " has made you an administrator!!" & vbCrLf, svr, True
620         Text "These powers have been given to you because " & RUser.RUserName & " trusts you with them" & vbCrLf, svr, True
630         Text "Abusing these powers can see you get kicked, or even banned from NChat!" & vbCrLf, svr, True
640         OldUsername = UserName
650         UserName = Replace(UserName, " [A]", "")
660         UserName = UserName & " [A]"
670         Broadcast "chuø" & OldUsername & MyIcon
680     End If

690   Case "con"    ' Connect a user
700     Log "Connect: " & RUser.RUserName & vbCrLf, 32768
710     NewUser = RUser.RUserName
720     DoEvents

730     Status CaptionPrefix & RUser.RUserName & " has entered the conversation..." & EndWindow

        ' Not yours? THen text it

740     If RUser.RIPAddress <> frmMain.sckUDP.LocalIP Then Text RUser.RUserName & " has entered the conversation..." & vbCrLf, "con", True, False, False, 2, "Center"

750     DoEvents
760     DoEvents

        ' Send your username to the new joiner.
        ' Send it to everyone else just in case they don't have it
770     If NewUser <> UserName Then
            'sckUDP.RemoteHost = sckUDP.RemoteHostIP

780         Broadcast "usrø" & MyIcon & "ø" & frmMain.sckUDP.LocalIP & "ø" & CreatedRoom
790     End If

        ' Add name to list
800     If UserName <> RUser.RUserName And RUser.RIPAddress <> frmMain.sckUDP.LocalIP Then
810         AddUser RUser.RUserName, Result(2), RUser.RIPAddress
820     End If

830     If frmMain.mnuDev.Visible = True And WelcomeMsg > "" Then
            ' If we don't replace WelcomeMSg with OWM, then
            ' things like +newuser+ can become stuck
840         OWM = WelcomeMsg
850         Broadcast WelcomeMsg
860         WelcomeMsg = OWM
870     End If

880   Case "isr"
        ' Are you a real user?
890     Log "Is Real Check: " & Result(2) & vbCrLf, 32768
900     If Result(2) = UserName Then sckUDP.SendData "N©H@-|-øsvrø" & UserName & "'s account is active!"

        ' The code to add users to your list
910   Case "usr"
920     UBR = UBound(Result) - 3
930     If Result(UBR) = "True" Then RoomHost = Result(2)
940     If RUser.RUserName = "" Or RUser.RUserName = UserName Then Exit Sub

        ' Don't add it if it's already there
950     If FindUser(RUser.RUserName) > 0 Then Exit Sub

        'If FindIcon(Result(3)) = -1 Then
        'Broadcast "rdlø" & Result(2)
        'Do Until FindIcon(Result(3)) <> -1
        'DoEvents
        'DoEvents
        'Loop
        'End If

960     AddUser RUser.RUserName, Result(2), RUser.RIPAddress


970   Case "pban"
980     Log Result(2) & " has been permanently banned from nchat by " & RUser.RUserName & vbCrLf, vbRed
990     If Result(2) = UserName Then
1000        MsgBox "YOU HAVE BEEN PERMANENTLY BANNED FROM NCHAT. " & UCase(RUser.RUserName) & " HAS DECIDED THAT YOU ARE UNFIT TO PARTICIPATE IN FURTHER NCHAT DISCUSSIONS. PLEASE CONTACT GRAYDA TO DISCUSS YOUR RETURN TO NCHAT", vbCritical, "PERMANENT BAN"
1010        Ban = True
1020  frmMain.MenuStuff 0
1030    End If

        ' When you create a new room, you become Room Admin
        ' This allows you to add new room admins
1040  Case "ad1"
1050    Log "Add Room Admin: " & Result(2) & " added by " & RUser.RUserName & vbCrLf, vbRed
1060    If Result(2) = UserName Then

1070        Text RUser.RUserName & " has made you a room Administrator" & vbCrLf, svr, True
1080        Text "While a room administrator doesn't have as much power as a full" & vbCrLf, svr, True
1090        Text "Administrator, they hold a great deal of power. Please use this power carefully" & vbCrLf, svr, True
1100        Text "Or face being kicked or even banned from NChat!" & vbCrLf, svr, True
1110        OldUsername = UserName
1120        UserName = Replace(UserName, " [RA]", "")
1130        UserName = UserName & " [RA]"
1140        Broadcast "chuø" & OldUsername & "ø" & UserName

1150        Status CaptionPrefix & " - You are now a Room Administrator" & EndWindow
1160    End If

        ' Chucks a room admin out to the curb :P
1170  Case "ad2"
1180    Log "Rem Room Admin: " & Result(2) & " removed by " & RUser.RUserName & vbCrLf, vbRed
1190    If Result(2) = UserName Then


1200        OldUsername = UserName
1210        UserName = Replace(UserName, " [RA]", "")
1220        Broadcast "chuø" & OldUsername & "ø" & UserName

1230        Text "Your room Administrator rights have been removed by " & RUser.RUserName & vbCrLf, svr, True
1240        Status CaptionPrefix & " - Your admin rights have been removed" & EndWindow
1250    End If

1260  Case "pm1"
        ' Concerned about privacy? Comment out this line.
        ' It stops people from listening in to Private
        ' conversations
        ' It's just a way for me to test Private Messages.
1270    Log "PM: " & RUser.RUserName & " (to " & Result(2) & "): " & Result(3) & vbCrLf, vbBlack, True

        ' Is it for you?
1280    If Result(3) = "" Then Exit Sub
1290    LastFrom = UserName
1300    If Result(2) = UserName Then    'And Result(4) <> Username Then
1310        Randomize

            ' If this is false, then don't show the box.
            ' This only changes after you recieve the
            ' first message.
1320        If NewMessage = False And frmMain.Visible = False Then
1330            Tray.Box RUser.RUserName & " has sent you a private message! Double click this icon to read it!", "New Private Message"
1340            NewMessage = True
1350        End If

1360        For I = 1 To lstIgnore.ListItems.Count
                'If lstIgnore.ListItems.Item(i) = "" Then Exit For
1370            If lstIgnore.ListItems.Item(I) = Result(4) Then
1380                Text RUser.RUserName & " (Who is on your ignore list), tried to send you a message", svr, True
1390                DoAutoBot sData
1400                Exit Sub
1410            End If
1420        Next I

1430        I = FindChatWindow(Result(4))
1440        If I > 0 Then
1450            If CW(I).Visible = True Then

1460                Txt2 RUser.RUserName & " ::  " & Result(3) & vbCrLf, act, Int(I)
1470                DoAutoBot sData

1480            End If
1490        End If


1500        If ListView1.Visible = False Then Text "You have a new Private Message from " & RUser.RUserName & vbCrLf, svr, True
1510        ListView1.ListItems.Add , "R…" & Result(3) & "…" & Int(Rnd * 5000), RUser.RUserName, , ImageList1.ListImages.Item(ImageList1.ListImages.Count - 2).Key
            ' Got an Away Message set? Display it then

1520        If AwayMSG <> "" And AwayMessage = True And Result(UBound(Result)) = "AutoAway" Then
1530            Broadcast "pm1+d+" & RUser.RUserName & "+d+" & AwayMSG & "+d+AutoAway"
1540        End If

1550    End If

        ' Disconnects a user from NChat and removes their name from
        ' the list of users
1560  Case "dis"
1570    Log "Disconnect: " & RUser.RUserName & vbCrLf, vbRed
1580    Text RUser.RUserName & " has left the room!" & vbCrLf, "dis", True, , , , "Center"
1590    Status CaptionPrefix & RUser.RUserName & " has left the room!!" & EndWindow

        ' Scroll through the list and removes the user in question
1600    RemoveUser RUser.RUserName
1610    DoEvents

        ' Gets rid of an admin's rights
        ' Unless you are a "True" admin
1620  Case "rem"
1630    Log "Kill Admin: " & Result(2) & " killed by " & RUser.RUserName & vbCrLf, vbRed
1640    If Result(2) = UserName And TrueAdmin = False Then
1650        frmMain.mnuDev.Visible = False
1660        frmMain.mnuUE.Visible = False
1670        OldUsername = UserName
1680        UserName = Replace(UserName, " [A]", "")
1690        Broadcast "chuø" & OldUsername & "ø" & MyIcon
1700        Text RUser.RUserName & " has taken away your admin rights!" & vbCrLf, "svr", True

1710    ElseIf Result(2) = UserName And TrueAdmin = True Then
1720        Text RUser.RUserName & " tried to take your admin rights! What a rat!!" & vbCrLf, "svr", True
1730    End If

1740  Case "move"
        ' This case deals with the whiteboard.
        ' The line colour, location, and size are sent in one
        ' packet, to save some network bandwidth (pfft. Yeah right. To draw
        ' a 3cm line, takes about 50 move commands...). The board
        ' should really have it's own winsock control, and be private
        ' between the 2 people, but I can't be stuffed...

        'Log "Move Cursor: " & Result(2) & " from " & Result(3) & " X: " & Result(4) & " Y: " & Result(5) & vbCrLf
1750    If Result(2) = UserName Then
1760        For I = LBound(CW) To UBound(CW)

1770            If CW(I).Tag = RUser.RUserName Then
1780                With CW(I)
1790                    .Picture1.DrawWidth = Result(6)
1800                    .Picture1.Enabled = False
1810                    .Picture1.Line (lastX, lastY)-(Result(3), Result(4)), Result(5)

1820                    lastX = Result(3)
1830                    lastY = Result(4)
1840                    .Picture1.Enabled = True
1850                End With
1860            End If

1870        Next I
1880    End If


1890  Case "fill"
        'Log "Fill " & Result(2) & "'s Picture Box (From " & Result(3) & ") at points X:" & Result(4) & " Y:" & Result(5) & " with the colour " & Result(6), vbRed, True

1900    If Result(2) = UserName Then
1910        For I = LBound(CW) To UBound(CW)

1920            If CW(I).Tag = RUser.RUserName Then
                    'ExtFloodFill CW(i).Picture1.hDC, Result(4), Result(5), Result(6), 1
1930                CW(I).Picture1.FillColor = Result(6)
1940                ExtFloodFill CW(I).Picture1.hdc, Result(4), Result(5), CW(I).Picture1.Point(Result(4), Result(5)), 1
1950            End If
1960        Next I
1970    End If



1980  Case "Clear"
        'Log Result(2) & "'s Whiteboard has been cleared by " & Result(3)
1990    If Result(2) = UserName Then
2000        For I = LBound(CW) To UBound(CW)

2010            If CW(I).Tag = RUser.RUserName Then
2020                CW(I).Picture1.Cls
2030                Exit Sub
2040            End If


2050        Next I
2060    End If

        ' Change your username and your Icon on all user lists
2070  Case "chu"
2080    Log "Change Username: " & Result(2) & " to " & Result(3) & vbCrLf, 32768

2090    TheirIndex = FindUser(Result(2))
        ' Can't change what's not there
2100    If TheirIndex = 0 Then Exit Sub

2110    frmMain.List1.ListItems.Item(TheirIndex).Text = Result(3)
2120    frmMain.List1.ListItems.Item(TheirIndex).SmallIcon = frmMain.ImageList1.ListImages.Item(Result(3)).Key

        ' Makes a really big heading to notify of events, crush people
        ' or just for fun
2130  Case "hea"
2140    Log "Heading: " & Result(2) & " (Sent by: " & RUser.RUserName & ")" & vbCrLf, vbBlue
2150    Text Result(2) & vbCrLf, "Heading", True, False, False, 8, "Center"
        '2440  Text "" & vbCrLf, "Msg", False, False, False, iChat.Font.Size


        ' when you purchase a kick user item from the store,
        ' it kicks someone and says "<USERNAME> has kicked you from the
        ' room", rather than saying that the admin did it
2160  Case "kun"
2170    Log RUser.RUserName & " kicked " & Result(2) & vbCrLf, vbRed
2180    If Result(2) = UserName Or Result(2) = sckUDP.LocalIP And TrueAdmin = False Then

2190        MsgBox RUser.RUserName & " has kicked you from the NChat Chatrooms", vbCritical, "Kicked by User"
2200        mnuDev.Visible = False
2210        If TrueAdmin = False Then
2220            frmMain.MenuStuff 0
2230            Broadcast "disø" & UserName
2240        Else
2250            Text RUser.RUserName & " tried to kick you from NChat, what a rat!!" & vbCrLf, "svr", True
2260        End If
2270    End If

        ' Another kind of kick, this time from the admin
2280  Case "ksv"
2290    Log RUser.RUserName & " (Admin) kicked " & Result(2) & vbCrLf, vbRed
2300    If Result(2) = UserName Or Result(2) = frmMain.sckUDP.LocalIP And TrueAdmin = False Then

            'mnuDev.Visible = False

2310        If TrueAdmin = False Then
2320            Broadcast "disø" & UserName

2330            Broadcast "svrø+username+ has been kicked from NChat"
2340            MsgBox "The administrator (" & RUser.RUserName & ") has kicked you from the room. Please correct your behaviour before re-connecting.", vbCritical, "Kicked by Administrator"

2350            frmMain.MenuStuff 0
2360        Else
2370            Text RUser.RUserName & " tried to kick you from NChat, what a rat!!", "svr", True
2380        End If

2390    End If

2400  Case "snd"
2410    Log RUser.RUserName & " sent " & Result(2) & " " & Result(3) & " NCredits" & vbCrLf, ThatsGood
        ' Sends NCredits to a user
        ' Result(2) = Username
        ' Result(3) = How Many NCredits
        ' Result(4) = Who are the NCredits from?

2420    If Result(2) = UserName Then

2430        NCredits = NCredits + Result(3)
2440        Text RUser.RUserName & " has given you " & Result(3) & " NCredits!" & vbCrLf, "ThatsGood", True
2450        Status CaptionPrefix & " - Someone has given you " & Result(3) & " NCredits!" & EndWindow
2460    End If

        ' Force kick someone. Only use if True Admins are misbehaving. Not a documented feature
2470  Case "force"
2480    If Result(2) = UserName Then
2490        MsgBox "Because of your actions, you have been kicked from NChat. Please correct your behaviour before re-entering!", vbCritical, "Force Kicked"
2500        mnuDev.Visible = False
2510        TrueAdmin = False
2520        frmMain.MenuStuff 0
2530    End If

        ' Prints a user's statistics
2540  Case "pip"
2550    Log Result(2) & "'s Stats Requested by " & RUser.RUserName & vbCrLf, 32768
2560    If RUser.RUserName = UserName Then
2570        sckUDP.SendData "N©H@-|-øsvrø" & UserName & "'s statistics (" & Time & "):+newline++newline+IP: " & sckUDP.LocalIP & "+newline+NCredits: " & NCredits & "+newline+Last Message from: " & sckUDP.RemoteHostIP & "+newline+Time on NChat: " & NChatTime & " seconds+newline+Core Version: " & App.Major & "/" & App.Minor & "/" & App.Revision & "+newline+Smileys on?: " & Smiley & "+newline+Real Username: " & GetUserName & "+newline+Admin: " & mnuDev.Visible & "+newline+Swearing on?: " & Swearing & "+newline+Total Icons: " & TotalIcons & "+newline+Profile: " & Profile & "+newline+On Computer: " & Environ("Computername")
2580        Status CaptionPrefix & " - Your statistics requested" & EndWindow
2590    End If

        ' Change your userlist icon. If it can't find an icon, then "Default" is used
2600  Case "chi"
2610    Log RUser.RUserName & " changed icon to index " & Result(2) & vbCrLf, 32768
2620    If FindIcon(Result(2)) = 0 Then
2630        If FindUser(RUser.RUserName) > 0 Then frmMain.List1.ListItems(FindUser(RUser.RUserName)).SmallIcon = frmMain.ImageList1.ListImages(FindIcon("Default")).Key
2640    End If

2650    If FindUser(RUser.RUserName) > 0 Then frmMain.List1.ListItems(FindUser(RUser.RUserName)).SmallIcon = frmMain.ImageList1.ListImages(FindIcon(Result(2))).Key

        ' Redirect a user to another room
2660  Case "red"
2670    Log Result(3) & " redirected to: " & Result(2) & " (" & Result(4) & ") by " & RUser.RUserName & vbCrLf, vbRed
2680    If IsInternet = True And Result(3) = UserName Then
            ' Click the disconnect menu
2690        frmMain.MenuStuff 1
2700        MsgBox "You have been redircted from the NChat server. You are now connected to NChat via the NETWORK. To re-connect to the NChat server, open the Main Menu", vbCritical, "You have been redirected"
2710        Exit Sub
2720    End If

2730    If Result(3) = UserName Then
2740        Text "You have been redirected to room #" & Result(2) & " (" & Result(3) & ")" & vbCrLf, "ThatsBad", True
'2750        Broadcast "disø" & UserName
2751        DoEvents
DoEvents
2760        NewRoom Result(2), Result(4)
2780    End If

        ' This command designates a NEW room host to run custom-made rooms.
2790  Case "nrh"
2800    Log "New Room Host has been selected: " & Result(4) & vbCrLf, vbBlack, True
2810    If UserName <> Result(4) Then Exit Sub
2820    Text "You have been selected as the new Room Host by " & RUser.RUserName & ". This means that you can run the current room, kick people, change the welcome message and more. " & RUser.RUserName & " has selected you because they trust you to run their room while their away.", "svr", True, True

2830    CreatedRoom = True
2840    RoomName = Result(2)
2850    RoomID = Result(3)
        '2780  WelcomeMsg = Result(5)
2860    Description = Result(5)
2870    Password = Result(6)

        ' Lets me know if people are ghosting others >:D
2880  Case "fak"
2890    Log "GHOST: " & Result(2) & " - " & Result(3) & " (" & RUser.RUserName & ")" & vbCrLf, vbRed
2900    If FileObj.FileExists(Left(Profile, Len(Profile) - Len("index.htm")) & "messages.htm") = True Then
2910        Open Left(Profile, Len(Profile) - Len("index.htm")) & "messages.htm" For Input As #1
2920        Text Input(LOF(1), 1)
2930        Close #1
2940    End If
2950    Text StartMSG & " " & Result(2) & " " & EndMSG & " " & Result(3) & vbCrLf, "Msg"

        ' Changes the caption to say the new message
2960    Status CaptionPrefix & Result(2) & " - " & Result(3) & EndWindow

2970  Case "fakeu"
2980  Log "Fake user joined: " & Result(2) & " with IP " & Result(3), vbRed, True
2990  Text Result(2) & " has entered the conversation..." & vbCrLf, "con", True, False, False, 2, "Center"
3000  AddUser Result(2), "default", Result(3)

3010  Case Else
        ' Not any of the above? Treat it as unknown or
        ' Raw Data message
3020    If Result(1) = UserName Or Result(0) = "room" Then Exit Sub

3030    Text Result(0) & vbCrLf, "Msg"
3040    Log "Unknown Command / Raw Data: " & Result(0) & vbCrLf, vbBlue
3050  End Select

3060  frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture
3070  frmMain.iChat.Document.body.scrolltop = CLng(Len(frmMain.iChat.Document.body.innerHTML)) * 100


          ' Scans through the latest batch of text for
          ' a smiley :) :evil

3080  StartAt = 1

          ' Let our NChat Bot do it's magic
3090  DoAutoBot sData

          ' Resets the RemoteHost
          ' So all messages can reach their destinations
3100  If IsInternet = False Then frmMain.sckUDP.RemoteHost = Address

3110  Exit Sub

ErrorH:

3120  If Err.Number > 0 Then
3130  MsgBox "There was an error with NChat. Some data may not have been sent / recieved. NChat will NOT close, but will continue. Please contact Grayda at: firestorm_visual@hotmail.com or at www.nchat.tk and quote the following:" & vbCrLf & vbCrLf & "Line Number: " & Erl & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Error!"
3140  End If

End Sub

Public Sub ParseRoomData(Data As String)
    ' The SHA String that is included with with the data header
 Dim sckData As String
 
    Dim RoomResult() As String
    On Error Resume Next
       
    RoomResult = SplitVB5(sckData, "ø")
    
    frmLog.List1.AddItem sckData
    Select Case RoomResult(0)

    Case "lst"
        If CreatedRoom = True Then
            RoomBroadcast "roomø" & "R" & frmMain.sckUDP.LocalPort & "/ù/" & Encode(Password, "FireStormInc") & "ø" & RoomName & "ø" & MyIcon & "ø"
        End If

    Case "room"

        For I = 1 To frmRooms.List1.Nodes.Count
            If frmRooms.List1.Nodes.Item(I).Text = RoomResult(2) Then Exit Sub
        Next I

        If frmRooms.Visible = True Then frmRooms.List1.Nodes.Add "C", 4, RoomResult(1), RoomResult(2), ImageList1.ListImages.Item(FindIcon(RoomResult(3))).Key
        'frmRooms.List1.Nodes(frmRooms.List1.Nodes.Count - 1).Tag = Result(4)

    Case "info"
        If RoomResult(1) = RoomName And CreatedRoom = True Then
            If Password = "" Then
                RoomBroadcast "rriø" & RoomResult(2) & "øRoom Name: " & RoomName & "|||| " & Description & "||||Host: " & UserName & "||Using Version: " & App.Major & "." & App.Minor & "." & App.Revision & "||# of people in room: " & frmMain.List1.ListItems.Count & "||Room has been running for: " & RoomTime & " seconds||RoomCreate Software v: 2.1.0||Last person to enter: " & NewUser & "||NChat Room ID: " & sckUDP.RemotePort & "||Last message from: " & LastFrom
                Text RoomResult(1) & " has requested your room info!" & vbCrLf, svr, True
            ElseIf Password > "" Then
                RoomBroadcast "rriø" & RoomResult(2) & "øInformation for this room is currently not available, because it has been marked private"
                Text RoomResult(1) & " attempted to request your room info!!" & vbCrLf, svr, True
            End If
        End If

    Case "rri"
        If RoomResult(1) = UserName Then MsgBox Replace(RoomResult(2), "||", vbCrLf), vbInformation, "Room Information"

    Case "com"
        Text RoomResult(1) & vbCrLf, svr, True

    Case Else
        Exit Sub
    End Select

End Sub
