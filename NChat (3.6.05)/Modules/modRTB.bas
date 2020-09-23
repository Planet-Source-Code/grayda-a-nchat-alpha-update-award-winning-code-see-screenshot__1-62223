Attribute VB_Name = "modText"
Option Compare Text

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub Text(Text As String, Optional Colour As String, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As String, Optional Font As String, Optional CheckSmileys As String)
    On Error Resume Next
    Dim OurSource As String
    Dim TopHalf As String
    Dim BottomHalf As String
    
    ' OK, this is the NEW AND IMPROVED Text sub, which allows you to insert text
    ' into the HTML Document (frmMain.iChat). Let's step through it, coz it can
    ' get messy here

    ' Nothing to "text"? Then Don't!
    If Text = "" Then Exit Sub
    ' Should we check for smileys? If blank, then assume we will
    If CheckSmileys = "" Then CheckSmileys = "True"
    
    ' Now to get the text BEFORE and AFTER the <NChat_HTML> tags, so
    ' we can insert data in the middle of the HTML document, so we won't get
    ' messed up web-pages
    OurSource = frmMain.iChat.Document.documentElement.innerHTML
    
    Tmp = InStr(1, OurSource, "<NChat_HTML>", vbTextCompare)
    'MsgBox Len(OurSource) & " " & Tmp
    TopHalf = Left(OurSource, Tmp - 1)
    BottomHalf = Mid(OurSource, Tmp, Len(OurSource))
        
      
    ' Replace shortcuts with their values. Say for example you type
    ' +username+, it will be REPLACEd with your current username
    Text = Replace(Text, "+username+", UserName)
    Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
    Text = Replace(Text, "+room+", RoomName & " (" & Room & ")")
    Text = Replace(Text, "+ncredits+", NCredits)
    Text = Replace(Text, "+newline+", vbLf)
    Text = Replace(Text, "+newuser+", NewUser)
    Text = Replace(Text, "+remoteip+", frmMain.sckUDP.RemoteHostIP)
    Text = Replace(Text, "+ntime+", NChatTime)
    Text = Replace(Text, "+ver+", Ver)
    Text = Replace(Text, "+time+", Format(Time, "HH:mm"))
    'Text = Replace(Text, "+lastfrom+", LastFrom)
    Text = Replace(Text, "+sender+", Result(UBound(Result) - 1))
    
    ' Room Stuff
    Text = Replace(Text, "+roomname+", RoomName)
    Text = Replace(Text, "+roomid+", Room)
    Text = Replace(Text, "+roomuser+", frmMain.List1.ListItems.Count)
    Text = Replace(Text, "+roomhost+", RoomHost)

    FindPos = InStr(1, Text, "+result")
    If FindPos > 0 Then Text = Replace(Text, "+result" & Mid(Text, FindPos + Len("+result"), 1) & "+", Result(Mid(Text, FindPos + Len("+result"), 1)))

    FindPos = InStr(1, Text, "+icon|")
    If FindPos > 0 Then
        'Text = Trim(Replace(Text, "+icon:" & Mid(Text, FindPos + Len("+icon:"), InStr(FindPos + Len("+icon:"), Text, "+")), "<img src='" & Mid(Text, FindPos + Len("+icon:"), InStr(FindPos + Len("+icon:"), Text, "+") - 2) & ".gif'>"))
        OurIcon = Mid(Text, FindPos + Len("+icon|"), InStr(FindPos + 1, Text, "+") - 2)
        OurIcon = Left(OurIcon, Len(OurIcon) - 2)

        Text = Replace(Text, OurIcon, "<img src='" & GetTempPath & "User Icons\" & OurIcon & ".gif'>")

    End If
    DoEvents
    DoEvents

    ' Temp is our original Text. Text is later cleared so we can add tags etc... Don't ask!
    Dim temp As String
    ' Set Temp to equal Text so we can clear text and write stuff in there. It's easier
    ' when it comes time to insert the original text into the tags
    temp = Text
    ' Replace vblf with <br> for some reason. Why isn't it vbcrlf? nm. It works OK
    temp = Replace(temp, vbLf, "<br>")
    ' Clear text so we can write tags in there.
    Text = ""
    ' First up is our alignment tag, if there is one. <p align="center"> for example
    If Alignment > "" Then Text = Text & "<p align=" & Chr(34) & Alignment & Chr(34) & ">"

    ' Next, open a BOLD tag
    If Bold = True Then Text = Text & "<strong>"
    ' Next, an Italic tag
    If Italic = True Then Text = Text & "<em>"
    ' Then an underline tag
    If Underline = True Then Text = Text & "<u>"
    ' The word 'colour' is very misleading. Colour actually refers to a class in our HTML file
    'If DoAsHTML = False And Colour > "" Then
    Text = Text & "<span class=" & Chr(34) & Colour & Chr(34) & ">"
    ' Open a font tag
    Text = Text & "<font"
    ' No font? Write the default (Arial), but if one is specified, use that
    If Font = "" Then
        Text = Text & " face=" & Chr(34) & "Arial" & Chr(34)
    Else
        Text = Text & " face=" & Chr(34) & Font & Chr(34)
    End If
    ' The size of our font. So far it looks like this: <font face="Arial" size="3"
    If Size > 0 Then
        Text = Text & " size=" & Chr(34) & Size & Chr(34)
    Else
        Size = "2"
    End If



    ' Close the Font tag
    Text = Text & ">" & temp & "</font>"

    ' Are smileys enabled? Then look for them
    If CheckSmileys = "True" And Smiley = True Then

        ' Smileys are now loaded from an external file. They won't work if deleted, and
        ' They aren't stored in the .RES file for size reasons. Smiley codes are now
        ' loaded from the .SMI File in the Smileys folder
        IniFile(3) = AppPath & "Smileys\Smiley.SMI"
        ' What smileys we have. Don't confuse this with the Boolean "Smiley", which lets us
        ' know if Smileys are on or off
        Dim Smileys() As String
        ' Our buffer to hold the smileys
        Dim szBuf As String, Length As Integer
        ' And extra big buffer because otherwise only 1/2 of the file is loaded
        szBuf = String$(25500, 0)
        ' Load all of our smiley names (eg. [:)] into our buffer
        Length = GetPrivateProfileSectionNames(szBuf, 25500, IniFile(3))
        szBuf = Left$(szBuf, Length)
        ' Assign our buffer to our array
        Smileys = Split(szBuf, vbNullChar)
        ' Then search for smileys
        For I = 0 To UBound(Smileys)
            If InStr(1, Text, Smileys(I), vbTextCompare) > 0 And Smileys(I) <> ":" Then
                Text = Replace(Text, Smileys(I), "<img alt=" & Chr(34) & ReadText(Smileys(I), "Caption", 3) & Chr(34) & " src=" & Chr(34) & AppPath & "Smileys/" & ReadText(Smileys(I), "Location", 3) & Chr(34) & ">")
            End If
        Next I

    End If

    ' Close the span tags, if they were open
    If Style > "" Then Text = Text & "</span>"
    ' Close the underline tag
    If Underline = True Then Text = Text & "</u>"
    ' Italics too
    If Italic = True Then Text = Text & "</em>"
    ' Bold tags too
    If Bold = True Then Text = Text & "</strong>"
    ' Finally, the </p> if the first one has been opened
    If Alignment > "" Then Text = Text & "</p>"

    ' A BR at the end to make it all neat
    If InStr(1, Text, vbCrLf) > 0 Then Text = Text & "<br>"

    ' FINALLY!! Add the code to the HTML file.
If TopHalf = "" Or BottomHalf = "" Then
        'MsgBox "This profile is corrupted. In the code for the profile, it's missing <NChat_HTML>. Open the profile in Notepad, and write in there (At the bottom should be OK): <NChat_HTML> (including the brackets)", vbCritical, "Profile corrupt!"
        frmMain.iChat.Document.body.innerHTML = frmMain.iChat.Document.body.innerHTML & Text
    Else
        frmMain.iChat.Document.body.innerHTML = TopHalf & Text & BottomHalf
End If

   ' frmMain.iChat.Document.body.innerHTML = frmMain.iChat.Document.body.innerHTML & Text
    DoEvents
    DoEvents
End Sub

Public Sub Txt2(Text As String, Colour As String, ByRef window As Integer)
' Writes text into a private chat window, hence
' the extra 'window' syntax
    Text = Replace(Text, "+username+", UserName)
    Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
    Text = Replace(Text, "+room+", RoomName)
    Text = Replace(Text, "+ncredits+", NCredits)

    With CW(window).Text1
        .SelStart = Len(.Text)
        .SelLength = Len(.Text)
        .SelColor = Colour
        .SelText = Text
        .SelLength = 0
    End With

End Sub


Public Sub Log(Text As String, Optional Colour As String, Optional Bold As Boolean)
' Custom Text commands
' Like the Text Sub in here, but much smaller

' Writes some stuff into our administrator's log
' (frmLog).

    With frmLog.RichTextBox1
        .SelStart = Len(.Text)
        .SelLength = Len(.Text)
        .SelColor = Colour
        .SelBold = Bold
        .SelText = Text
        .SelLength = 0

    End With

End Sub

Public Function DectoWebCol(lngColour As Long) As String
    Dim strColour As String
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    'Add leading zero's


    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    'Reverse the bgr string pairs to rgb
    DectoWebCol = "#" & Right$(strColour, 2) & _
                  Mid$(strColour, 3, 2) & _
                  Left$(strColour, 2)
End Function

Public Function GetLongRGB(ByVal pWebColor As String) As Long
'
'Get the long color for web color
'
    On Error GoTo Cerr
    If Mid(pWebColor, 1, 1) <> "#" Then pWebColor = GetColor(pWebColor)
    If Mid(pWebColor, 1, 1) = "#" Then pWebColor = Mid(pWebColor, 2)
    pWebColor = Mid(pWebColor, 5, 2) & Mid(pWebColor, 3, 2) & Mid(pWebColor, 1, 2)
    GetLongRGB = CLng("&H" & UCase(pWebColor))
    Exit Function
Cerr:
    GetLongRGB = -1
End Function
Public Function GetColor(ByVal pColor As String) As String
'
'Get the color from the Color set
'
    Dim lPos As Long
    On Error Resume Next
    lPos = InStr(1, ColorSet, " " & pColor & "(#", vbTextCompare)
    If lPos > 0 Then
        lPos = lPos + Len(" " & pColor & "(#")
        If lPos > 0 Then
            GetColor = Mid(ColorSet, lPos, 6)
        End If
    End If
End Function
