VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Browsing ..."
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh User's List of Files"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5280
      Top             =   3000
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden List"
      Visible         =   0   'False
      Begin VB.Menu mnuDownload 
         Caption         =   "Download this file"
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' Refresh the list
' lst (Person's files to request) (Your name)
'   Retrieves the list of files
List1.Clear
Broadcast "lstø" & Me.Tag & "ø" & UserName
End Sub

Private Sub Form_Load()
' Upon opening, call for the list of files
Broadcast "lstø" & Me.Tag & "ø" & UserName
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
' Right click menu
If Button = 2 And List1.List(List1.ListIndex) > "" Then PopupMenu mnuHidden

End Sub

Private Sub mnuDownload_Click()
' Download the file into the recieved files folder

Receiver.Show

File2Save = AppPath & "Recieved Files\" & List1.Text
' Put in file size thingy here (?)

'Send.Show
' Request the download. I may impliment an ignore list
' for downloaders, or something, but nah... :)
Broadcast "dwlø" & Me.Tag & "ø" & List1.Text & "ø" & UserName & "ø" & frmMain.sckUDP.LocalIP & "ø" & List1.Tag
DoEvents
DoEvents

End Sub

Private Sub Timer1_Timer()
' Just some caption stuff
Me.Caption = "Browsing " & Me.Tag & "'s Files (" & List1.Tag & ")"

End Sub

