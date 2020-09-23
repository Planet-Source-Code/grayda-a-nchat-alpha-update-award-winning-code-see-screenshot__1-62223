Attribute VB_Name = "modHyperlink"
Option Compare Text
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
 ' Character from under the cursor
Public Const EM_CHARFROMPOS& = &HD7
' Code to paste (Like ctrl+v)
Public Const WM_PASTE = &H302
    Public point As POINTAPI
    Public charpos As Long
    Public pos_start As Long
    Public pos_end As Long
    Public char As String
    Public Word As String

' Everything here is for the URL Detection
' for rich text boxes. These are URL Prefixes
' such as www, and the suffixes, like .com
Public htxt As String 'Stores the current mouse over text.. for HYPERLINK LANCHIN
Public Const sFrom3Left = "www*ftp*wais*news*telnet*prospero*nntp*gopher*file*htpps*http"
Public Const sFrom3Right2 = "to*cc*tv*ws*ms*jp*ro*tc*ph*dk*st*ac*gs*vu*vg*sh*kz*as*lt*de*us*ca*tk"
Public Const sFrom3Right3 = "frm*com*net*org*biz" 'Why I do this??? just to show u hehe...This is easy for you if u need to add
Public Const sFrom3Right4 = "info" 'makin it faster u know... O_o seperating it and combine
Public Const sFrom3Right5 = "co.za*co.nz*co.il*co.uk*co.vt*co.jp" 'later in dependin in the
Public Const sFrom3Right6 = "org.il*net.nz*org.uk*org.nz*com.ph*com.au"  'size [len]
Public LeftDat() As String
Public RightDat() As String

' Used to open hyperlinks and stuff
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Used to paste the smileys into NChat
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' PointAPI is the current position of our cursor
' Useful for detecting if our mouse resides over
' a control, or window. or something else
Public Type POINTAPI
    x As Long
    y As Long
End Type

' This string ends up with our hyperlink
Public Hyperlink As String

' Dunno what the wm_user const is for, but it's
' used everywhere these days :)
Private Const WM_USER = &H400
' Automatically detects our URL and underlines it
' All we need to do then is raise an event when
' we click on it (ie. Run our browser with our site)
Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)

Public Const ENM_LINK = 16738603
Public Const EM_GETEVENTMASK = WM_USER + 59
Public Const EM_SETEVENTMASK = WM_USER + 69

Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000


Public Function SetAutoURL4RTB(RTB As RichTextBox) As Long
' This function underlines our hyperlink using
' the default windows colours. (ie. Blue).
' having it underlined doesn't mean we can click
' on it yet, we need to work that out later :)
    Dim lngDat As Long

    
lngDat = SendMessageLong(RTB.hwnd, EM_AUTOURLDETECT, Abs(True), 0)
    SetAutoURL4RTB = lngDat
End Function

Public Function GetHyperlink(rch As RichTextBox, x As Single, y As Single) As String
' Our magic code that retrieves our hyperlink from
' a richtextbox. It's fast, clean and effective
' and it can also be expanded to include any
' new domains that may spring up
    Dim pt As POINTAPI
    Dim pos As Long
    Dim ch As String
    Dim txt As String
    Dim txtlen As Long
    Dim pos_start As Long
    Dim pos_end As Long
    
    ' convert mouse pos in pixels
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' position of character under cursor
    pos = SendMessage(rch.hwnd, &HD7, 0&, pt)
    If pos <= 0 Then
        Exit Function
    End If
    txt = rch.Text

    ' get start position of word under cursor
    For pos_start = pos To 1 Step -1
        If Mid$(txt, pos_start + 1, 1) = Chr(13) Then
            Exit Function
        End If
        ch = Mid$(txt, pos_start, 1)
        If ch = Chr(32) Or ch = vbCr Or ch = vbLf Or ch = vbNewLine Then Exit For
    Next pos_start
    pos_start = pos_start + 1

    ' get end position of word under cursor
    txtlen = Len(txt)
    For pos_end = pos To txtlen
        ch = Mid$(txt, pos_end, 1)
    If ch = Chr(32) Or ch = vbCr Then Exit For
    Next pos_end
    pos_end = pos_end - 1

    If pos_start <= pos_end Then _
        GetHyperlink = Mid$(txt, pos_start, pos_end - pos_start + 1)
End Function


