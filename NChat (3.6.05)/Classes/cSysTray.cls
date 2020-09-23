VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This code was pieced together from various tutorials
' I found across the internet. They were created
' by: Dave, OutKast, Mr. Mellow and KeyB, as well as
' about 4 without names
Private Declare Function GetLastError Lib "kernel32" () As Long
Private m_IconData As NOTIFYICONDATA

Private IconVisible As Boolean
Private ChangeOnUpdate As Boolean

Public Event NIError(ByVal ErrorNumber As Long)

Public PopUpMessage As String
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type


Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const NIF_INFO = &H10



Private Const MAX_TIP_LENGTH As Long = 64
Private Const WM_MOUSEMOVE = &H200

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' I LOVE THIS SUB!!
' It lets you have the funky balloon tip that comes
' up on windows XP machines, when you went to burn
' files, but didn't remove them after, then reset
' your machine, then got a coke, only to find that
' I am rambling on
Public Sub Box(Message As String, Title As String)
    With m_IconData
        .cbSize = Len(m_IconData)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE    'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmMain.Icon
        .szTip = Title & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = NIIF_WARNING
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Public Function Initialize(ByVal hwnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
    With m_IconData
        .cbSize = Len(m_IconData)
        .hIcon = hIcon
        .hwnd = hwnd
        .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
        .uCallbackMessage = uCallbackMessage
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uID = vbNull
    End With
    IconVisible = False
    ChangeOnUpdate = True
End Function

Public Function ShowIcon() As Boolean

    If Not IconVisible Then
        ShowIcon = Shell_NotifyIcon(NIM_ADD, m_IconData)
        If ShowIcon = False Then
            RaiseEvent NIError(GetLastError)
        Else
            IconVisible = True
        End If
    End If
End Function

Public Function HideIcon() As Boolean

    If IconVisible Then
        HideIcon = Shell_NotifyIcon(NIM_DELETE, m_IconData)
        If HideIcon = False Then
            RaiseEvent NIError(GetLastError)
        Else
            IconVisible = False
        End If
    End If
End Function

Public Property Let IconHandle(ByVal hIcon As Long)
    m_IconData.hIcon = hIcon
    If ChangeOnUpdate Then
        m_IconData.uFlags = NIF_ICON
        Update
        m_IconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or TTS_BALLOON
    End If
End Property

Public Property Let TText(ByVal sTip As String)
    m_IconData.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
    If ChangeOnUpdate Then
        m_IconData.uFlags = NIF_TIP
        Update
        m_IconData.uFlags = NIF_ICON Or NIF_MESSAGE Or TTS_BALLOON
    End If
End Property

Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
    m_IconData.uCallbackMessage = uCallbackMessage
    If ChangeOnUpdate Then
        m_IconData.uFlags = NIF_MESSAGE Or TTS_BALLOON
        Update
        m_IconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or TTS_BALLOON
    End If

End Property

Public Function Update() As Boolean
    If IconVisible Then
        Update = Shell_NotifyIcon(NIM_MODIFY, m_IconData)
        If Update = False Then
            RaiseEvent NIError(GetLastError)
        End If
    End If
End Function

Public Property Get IconHandle() As Long

    IconHandle = m_IconData.hIcon
End Property

Public Property Get TText() As String

    TText = Left(m_IconData.szTip, Len(m_IconData.szTip) - 1)
End Property

Public Property Get CallbackMessage() As Long
    CallbackMessage = m_IconData.uCallbackMessage
End Property

Public Property Let UpdateOnChange(bUpdate As Boolean)
    ChangeOnUpdate = bUpdate
End Property

Private Property Get UpdateOnChange() As Boolean
    UpdateOnChange = ChangeOnUpdate
End Property

Private Sub Class_Terminate()
    HideIcon
End Sub

Public Property Get Visible() As Boolean
    If IconVisible Then
        Visible = True
    End If
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    If bVisible Then
        ShowIcon
    Else
        HideIcon
    End If
End Property



