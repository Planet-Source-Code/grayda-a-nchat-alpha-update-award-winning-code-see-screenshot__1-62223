VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NChat Alpha Generation 2"
   ClientHeight    =   7125
   ClientLeft      =   -1395
   ClientTop       =   -780
   ClientWidth     =   4455
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15000
      Left            =   0
      Picture         =   "frmAbout.frx":1B7A
      ScaleHeight     =   15000
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   6840
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   840
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   2520
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OK this form is our about form. All controls
' are contained in a picture box, so we can scroll
' the whole thing without too much code.

' I know there is a scrollhdc API, but this is simpler
' and scrolls at a rate that stops most flickering

' Get the cursor position, so we can detect
' if it's within our form
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' This works out the position of our control
' on the screen. This works hand-in-hand with
' the functions above.
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



' Dim our R1 as a new rect, or you
' can get a byref error. I sat here for an hour
' trying to work out what the byref error was.
' Turns out I forgot to dim my rect \:)
Dim R1 As RECT
Dim P1 As POINTAPI

Private Sub Timer1_Timer()
' Slowly scrolls our credits
' Doesn't flicker on my computer, not sure about others
    Picture1.Top = Picture1.Top - 25
    If Picture1.Top < -Picture1.Height - Me.Height Then Picture1.Top = Me.Height
End Sub

Private Sub Timer2_Timer()
    Dim R1 As RECT
    Dim P1 As POINTAPI
    ' Get the location of our form
    GetWindowRect Me.hwnd, R1
    ' Get the location of our cursor
    GetCursorPos P1

    ' Is our cursor over our form? If so, then stop scrolling
    If P1.x < R1.Right And P1.x > R1.Left And P1.y < R1.Bottom And P1.y > R1.Top Then
        Timer1.Enabled = False
    Else
        Timer1.Enabled = True
    End If

End Sub
