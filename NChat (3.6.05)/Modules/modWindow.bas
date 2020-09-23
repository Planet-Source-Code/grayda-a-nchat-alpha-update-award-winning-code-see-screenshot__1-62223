Attribute VB_Name = "modWindow"
' Say you are talking to user2 and user3,
' Any message from user2 will only come into user2's
' window. Messages from user3 will go into it's
' respective window. If user4 joins in, a new window
' will be created for them
Public CW(1 To 100) As New frmChat

' Consts for Always On Top stuff
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

' This fades out the window when you exit NChat
' API calls and Constants from allapi.net's API List
Public Const AW_HIDE = &H10000    'Hides the window. By default, the window is shown.
Public Const AW_BLEND = &H80000    'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

' Sets the position of the window (Including Z Order)
' This code from allapi.net's API Guide
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

' All these colour constants are for profiles
' They pretty much describe what they colour

' Heading: Headings and large text
' Msg: Standard user messages
' Dis: User Disconnects
' Svr: Server (Purple) Messages
' Act: WinMX Actions (Emotes)
' Con: User connects
' ThatsGood: Good news
' ThatsBad: Bad news / Errors
Public Heading As String
Public Msg As String
Public dis As String
Public svr As String
Public act As String
Public con As String
Public ThatsGood As String
Public ThatsBad As String

' For drawing lines on our Whiteboard
Public Col As ColorConstants

' Caption prefix is the form's message. For example,
' If Grayda sends the message Hi, and my
' CaptionPrefix is: NChat [
' EndWindow is: ]
' Then my form caption will be:
' NChat [ Grayda - Hi ]
Public CaptionPrefix As String
Public EndWindow As String

' OldCap is the first part of the frmMain Caption
' (eg. Welcome to NChat - <Rest of the caption goes here>)
Public OldCap As String

' System tray stuff
Public Tray As New cSysTray

' This stuff is for displaying that really cool "Browse for Folder" window
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
' End "Browse for Folder" window declarations, const and types

Public Sub OnTop(window As Long)
'Set the window position to topmost
    SetWindowPos window, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
End Sub

Public Sub Status(Tip As String, Optional Icon As Long)
' The Status sub does many things:
' It sets frmMain's caption, sets the Tray's tip
' and if necessary, sets the tray's icon
    OldCap = frmMain.Caption

    If Tip > "" Then
        frmMain.Caption = Tip
        Tray.TText = Tip
        Tray.Update
    End If

    If Icon > 0 Then
        Tray.IconHandle = Icon
    End If

End Sub

Public Function ShowBox(ButtonText As String, WindowTitle As String) As String
' Showbox lets us do many things.
' for example, it lets us pick one username from the
' list of users, so we can kick them, print their info
' or anything that requires the end user to pick a name

    frmKick.Show
    frmKick.Command1.Caption = ButtonText
    frmKick.Caption = WindowTitle
    Do Until frmKick.Visible = False
        DoEvents
        DoEvents
    Loop
    ShowBox = SelUser
End Function

Public Function FindChatWindow(UserName As String)

    For I = LBound(CW) To UBound(CW)

        If CW(I).Tag = UserName Then
            FindChatWindow = I
            Exit Function
        End If

    Next I
    FindChatWindow = 0

End Function

Public Function FindFreeWindow() As Integer
    For I = LBound(CW) To UBound(CW)

        If CW(I).Tag = "" Then
            FindFreeWindow = I
            Exit Function
        End If

    Next I
    FindFreeWindow = 0
End Function

' Lets us Browse for a folder using the Windows API call
Public Function BrowseForFolder() As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = frmMain.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(AppPath, "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseForFolder = sPath

End Function

