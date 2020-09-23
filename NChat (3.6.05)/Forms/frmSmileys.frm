VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmSmileys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Smileys that you can use in NChat"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmSmileys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to download more smileys from www.solidinc.tk! (Internet Connection REQUIRED)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   8175
   End
   Begin SHDocVwCtl.WebBrowser Smiley 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmSmileys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click(Index As Integer)

End Sub

Private Sub Command1_Click()
    Smiley.Navigate "http://www.solidinc.aspfreeserver.com/nchat/index.php"
End Sub

Private Sub Form_Load()
'On Error Resume Next
    Dim AText As String
    temp = AppPath
    For I = 101 To 142
        AText = AText & "<img src=" & Chr(34) & temp & "smileys\" & "Smiley" & I & ".gif" & Chr(34) & ">" & LoadResString(I) & "<br>"
    Next I
    frmSmileys.Show
    DoEvents
    DoEvents

    Smiley.Document.body.innerHTML = AText

End Sub



