VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAutoBotOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Notch Control Panel"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   Icon            =   "frmAutoBotOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Idle Chat Options"
      TabPicture(0)   =   "frmAutoBotOptions.frx":1B7A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "Timer1"
      Tab(0).Control(4)=   "Check1"
      Tab(0).Control(5)=   "Text2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Phrase Learning"
      TabPicture(1)   =   "frmAutoBotOptions.frx":1B96
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Timer2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "NChat_Button1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "NChat_Button2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "RSS Newsfeeds"
      TabPicture(2)   =   "frmAutoBotOptions.frx":1BB2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(2)=   "Text1"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Scripting"
      TabPicture(3)   =   "frmAutoBotOptions.frx":1BCE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.CheckBox Check2 
         Caption         =   "Enable Scripting? (Recommended)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   13
         Top             =   1620
         Width           =   4095
      End
      Begin VB.CommandButton NChat_Button2 
         Caption         =   "Start Learning"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton NChat_Button1 
         Caption         =   "Stop Learning"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   6120
         Top             =   2460
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74040
         TabIndex        =   3
         Text            =   "10"
         Top             =   2100
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Allow Idle Chatter?"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   -69120
         Top             =   2100
      End
      Begin VB.Label Label8 
         Caption         =   $"frmAutoBotOptions.frx":1BEA
         Height          =   855
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label7 
         Caption         =   "Location of RSS File:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   $"frmAutoBotOptions.frx":1D5D
         Height          =   615
         Left            =   -74760
         TabIndex        =   11
         Top             =   780
         Width           =   6255
      End
      Begin VB.Label Label5 
         Caption         =   $"frmAutoBotOptions.frx":1E6B
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   6255
      End
      Begin VB.Label Label4 
         Caption         =   $"frmAutoBotOptions.frx":1F48
         Height          =   855
         Left            =   -74880
         TabIndex        =   9
         Top             =   780
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "New Phrases added to DB: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2460
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   -73080
         TabIndex        =   5
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Interval:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   2100
         Width           =   615
      End
   End
   Begin VB.CommandButton NChat_Button4 
      Caption         =   "Hide Window"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "frmAutoBotOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
' Timer1 handles our idle chatter,
' When Timer1_Timer is triggered,
' then Notch will speak a random phrase
    If Check1.Value = 1 Then
        Timer1.Enabled = True
        frmAutobot.List1.AddItem "(" & Time & ") Notch will now start speaking random phrases every " & Timer1.Interval / 10000 & " seconds."
    Else
        Timer1.Enabled = False
        frmAutobot.List1.AddItem "(" & Time & ") Notch has stopped speaking random phrases"
    End If


End Sub


Private Sub Check2_Click()
    AllowScripting = Not AllowScripting
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
End Sub

Private Sub NChat_Button1_Click()
    NotchLearning = False
    frmAutobot.List1.AddItem "(" & Time & ") Notch has stopped learning new phrases. He learnt " & NewWords & " new phrases!"

End Sub

Private Sub NChat_Button2_Click()
    NotchLearning = True
    frmAutobot.List1.AddItem "(" & Time & ") Notch is currently learning new phrases"


End Sub

Private Sub NChat_Button4_Click()
    Me.Visible = False

End Sub


Private Sub Timer1_Timer()
' This is our Idle Phrase broadcaster
    Dim Sects As Integer
    Dim ThePhrase As String

    If IniFile(2) = "" Then Exit Sub
    Sects = 1
    Do Until ReadText("IdlePhrase", "Phrase" & Sects, 2) = ""
        Sects = Sects + 1
    Loop
    ' Ensures 0 doesn't come up
    R = Int(Rnd * Sects) + 1

    ThePhrase = ReadText("IdlePhrase", "Phrase" & R, 2)
    ThePhrase = Replace(ThePhrase, "%rss%", GetRSSHeadline(frmAutoBotOptions.Text1.Text))
    frmAutobot.List1.AddItem "(" & Time & ") Notch has broadcasted IdlePhrase: " & R
    Broadcast ThePhrase
End Sub

Private Sub Timer2_Timer()
    Label1.Caption = "New Phrases added to DB: " & NewWords
End Sub
