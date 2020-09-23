VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Welcome to NChat Alpha - Waiting for incoming connections..."
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1920
      Top             =   2520
   End
   Begin VB.ComboBox txtSend 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   8055
   End
   Begin MSScriptControlCtl.ScriptControl scNotch 
      Left            =   3120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      Timeout         =   10000000
      UseSafeSubset   =   -1  'True
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "/action"
      Height          =   255
      Left            =   9480
      TabIndex        =   3
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send!!"
      Height          =   255
      Left            =   8400
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Tag             =   " "
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B7A
            Key             =   "Default"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F14
            Key             =   "Devil"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22AE
            Key             =   "Agent Smith"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2648
            Key             =   "Radioactive"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29E2
            Key             =   "Smiley"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CAD
            Key             =   "NChat"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59B7
            Key             =   "Star"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D51
            Key             =   "Lightning"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60EB
            Key             =   "Half-Life"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6485
            Key             =   "One Fingered Salute"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":681F
            Key             =   "GTAIII"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BB9
            Key             =   "Boy n Girl"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F6C
            Key             =   "Chinese"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":753A
            Key             =   "The Finger"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A3C
            Key             =   "Idiot"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E21
            Key             =   "Weed"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82BD
            Key             =   "Power"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8657
            Key             =   "Play"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89F1
            Key             =   "MOHA"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D8B
            Key             =   "Girl"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":913D
            Key             =   "Boy"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94EF
            Key             =   "Rammstein 1"
            Object.Tag             =   "Music"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A89
            Key             =   "Rammstein 2"
            Object.Tag             =   "Music"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E23
            Key             =   "Green Eye"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3BD
            Key             =   "Evanescence"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A957
            Key             =   "Evil 1"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B231
            Key             =   "Evil 2"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B7CB
            Key             =   "Nemo"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BD65
            Key             =   "Intruder"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C63F
            Key             =   "Gold Star"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF19
            Key             =   "Delta Force"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7F3
            Key             =   "Outkast"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E0CD
            Key             =   "Ozzy Osbourne"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9A7
            Key             =   "Gun"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F25D
            Key             =   "Halo"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF37
            Key             =   "Alert!!"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C11
            Key             =   "Birdman"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10FF8
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1204A
            Key             =   "Unavailable"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":123E4
            Key             =   "AFK"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckRooms 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   2222
      LocalPort       =   2222
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   7170
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3016
            MinWidth        =   882
            Text            =   "NChat - Disconnected!"
            TextSave        =   "NChat - Disconnected!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   423
            Picture         =   "frmMain.frx":1277E
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2040
      Top             =   720
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   1440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save NChat Log"
      Filter          =   "Text Files (*.TXT)|*.txt|RTF File (*.rtf)|*.RTF|All Files (*.*)|*.*"
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   1113
      LocalPort       =   1113
   End
   Begin MSComctlLib.ListView List1 
      Height          =   6525
      Left            =   8400
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   11509
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   3413
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser iChat 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8175
      ExtentX         =   14420
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   6525
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   11509
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstIgnore 
      Height          =   6525
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   11509
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Image picGreen 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":12AD0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Main Menu"
      Begin VB.Menu mnuDownload 
         Caption         =   "Download NChat-Packs"
      End
      Begin VB.Menu mnuConnection 
         Caption         =   "Connection"
         Begin VB.Menu mnuConnect 
            Caption         =   "Connect to NChat Server"
         End
         Begin VB.Menu hrDL 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDisconnect 
            Caption         =   "Disconnect from Server"
         End
      End
      Begin VB.Menu MMHR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadProfile 
         Caption         =   "Load Profile"
         Shortcut        =   ^L
      End
      Begin VB.Menu hr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveChat 
         Caption         =   "Save Chat Log"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Chat Text"
      End
      Begin VB.Menu hr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "NChat Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu hr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit NChat"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuNCredits 
      Caption         =   "NCredits"
      Begin VB.Menu mnuBalance 
         Caption         =   "NCredits Balance"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuStore 
         Caption         =   "NChat Store"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuUL 
         Caption         =   "List of people in the room"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPM 
         Caption         =   "Private Messages"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuIgnore2 
         Caption         =   "Ignore List"
      End
   End
   Begin VB.Menu mnuChatRooms 
      Caption         =   "Chat Rooms"
      Begin VB.Menu mnuJoinCustom 
         Caption         =   "List all rooms"
      End
      Begin VB.Menu mnuCreateRoom 
         Caption         =   "Create your own room"
      End
      Begin VB.Menu hr5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChangeRoomInfo 
         Caption         =   "Change room info"
         Visible         =   0   'False
      End
      Begin VB.Menu hr12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAKick 
         Caption         =   "Kick user"
      End
      Begin VB.Menu mnuRAWelcome 
         Caption         =   "Welcome message"
      End
      Begin VB.Menu mnuRAPassword 
         Caption         =   "Change room password"
      End
      Begin VB.Menu mnuRAServerMSG 
         Caption         =   "Send Server (Purple) Message"
      End
   End
   Begin VB.Menu mnuDev 
      Caption         =   "Admin Menu"
      Begin VB.Menu mnuAdminWindow 
         Caption         =   "NChat Server Log"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuBL 
         Caption         =   "Broadcast / Loopback"
      End
      Begin VB.Menu mnuRawData 
         Caption         =   "Raw Data"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAutoBot 
         Caption         =   "Start / Stop Bot"
      End
      Begin VB.Menu mnuChangeNCredits 
         Caption         =   "Change your NCredits"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuChatGoodies 
         Caption         =   "Chatroom Goodies"
         Begin VB.Menu mnuFake 
            Caption         =   "Fake Users"
            Begin VB.Menu mnuInsertFake 
               Caption         =   "Insert Fake User"
               Shortcut        =   ^I
            End
            Begin VB.Menu mnuRemFake 
               Caption         =   "Remove Fake User"
               Shortcut        =   ^D
            End
         End
         Begin VB.Menu mnuWmsg 
            Caption         =   "Welcome Message"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuDoHeading 
            Caption         =   "Create a heading"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuSendAll 
            Caption         =   "Send a message to all rooms"
         End
         Begin VB.Menu mnuNewRoom 
            Caption         =   "New Room with set ID"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTextHelp 
         Caption         =   "Display Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSmileys 
         Caption         =   "Smileys you can use"
      End
      Begin VB.Menu mnuTipofTheDay 
         Caption         =   "Tip of the day"
      End
      Begin VB.Menu hrA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About NChat"
      End
   End
   Begin VB.Menu mnuUserList 
      Caption         =   "User List"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ul1 
         Caption         =   "Send >> a Private Message"
      End
      Begin VB.Menu ul2 
         Caption         =   "Send >> some NCredits"
      End
      Begin VB.Menu mnuUEIgnore 
         Caption         =   "Ignore >>"
      End
      Begin VB.Menu mnuUE 
         Caption         =   "Admin"
         Visible         =   0   'False
         Begin VB.Menu mnuUEGhost 
            Caption         =   "Ghost"
         End
         Begin VB.Menu mnuUEKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu mnuUEAdmin 
            Caption         =   "Make Admin"
         End
         Begin VB.Menu mnuUERemAdmin 
            Caption         =   "Kill Admin"
         End
         Begin VB.Menu mnuUERedirect 
            Caption         =   "Redirect"
         End
         Begin VB.Menu mnuUEPIP 
            Caption         =   "Print Info"
         End
         Begin VB.Menu hr6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRealUser 
            Caption         =   "Is this user real?"
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Information Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAvailable 
         Caption         =   "I am Available for chat"
      End
      Begin VB.Menu mnuAFK 
         Caption         =   "I am Away from keyboard"
      End
      Begin VB.Menu mnuUnAvailable 
         Caption         =   "I am Unavailable for chat"
      End
      Begin VB.Menu hr7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAwayMSG 
         Caption         =   "Set Custom Away Message"
      End
      Begin VB.Menu hr8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore List"
      End
      Begin VB.Menu hr9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUECHI 
         Caption         =   "Change my User Icon"
      End
      Begin VB.Menu mnuUECHU 
         Caption         =   "Change my Username"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "NChat Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuTray_Show 
         Caption         =   "Show NChat"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Completely Hide NChat"
      End
      Begin VB.Menu hr10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray_Available 
         Caption         =   "I am Available for chat"
      End
      Begin VB.Menu mnuTray_AFK 
         Caption         =   "I am Away From Keyboard"
      End
      Begin VB.Menu mnuTray_Unavailable 
         Caption         =   "I am Unavailable for chat"
      End
      Begin VB.Menu hr11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray_Quit 
         Caption         =   "Quit NChat"
      End
   End
   Begin VB.Menu mnuPopup_Ignore 
      Caption         =   "Ignore Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteIgnore 
         Caption         =   "Delete Ignore"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Solid Inc. Media Productions Presents:
'       NChat Alpha - Your Chat in a box!

' Thankyou for downloading the source code to NChat, my
' major Visual Basic programming task spanning the last
' 3 years. This program was originally written to allow people"msg"

' at my school to communicate throughout the 3 computer rooms and
' the library. It was also written to show people that chat projects
' don't have to end at 'emoticons' and 'RTF Support', that AI has a
' place in 'across the room' chat programs, and that using a points
' system can bring much needed fun to an otherwise dull program

' So please enjoy this code, and if you like it, please comment
' and vote. Every vote encourages me to make this chat program
' THE MOST FEATURE PACKED CHAT ON PLANET SOURCE CODE!!

' -Grayda

' This code allows incoming data to be case insensitive,
' meaning that USERNAME and uSeRnAmE are the same.
' In fact, ANY string is now case insensitive!
Option Compare Text

' Your Away Message, that is sent to everyone
' who tries to contact you via Private Messages
' when your status is set to anything other than
' available (ie. AFK, Unavailable)
Dim AwayMSG As String

' OldData was the last message to be recieved through
' the winsock. It exists, so people cannot rack
' up NCredits by spamming the room. You can still
' send lots of messages, but you won't recieve
' NCredits for them
Dim OldData As String

' If Ban = True, then when you exit NChat,
' you are banned forever, or until you
' delete your settings file and start again
Dim Ban As Boolean

Private Sub cmdAction_Click()
' An action lets you act out certain things.
' the action is replaced by your username,
' so /action screams becomes <USERNAME> screams

' This is widely used in IRC chat rooms and WinMX chat rooms
    txtSend.Text = "/action " & txtSend.Text
End Sub

Private Sub cmdSend_Click()
' The Send button at the bottom right of frmMain

' This sub has been completely re-written, because
' it had a small security risk in there. People
' could get 1 NCredits each time for flooding the
' room with /action messages

' Stop Empty-Messages from being sent
    If Trim(txtSend.Text) = "" Then Exit Sub
    ' This code ensures that you can't raise errors
    ' by trying to chat on an 'unconnected' sock
    If SB1.Panels(1).Text = "NChat - Disconnected!" Then
        MsgBox "There was an error sending your message. It appears that you are not connected to the NChat server. Check the little light in the corner. If it is red, then please wait 30 seconds, or restart NChat. But if it is green, then try and send the message again", vbCritical, "Not Connected"
        Exit Sub
    End If

    ' Oh yeah, we need to remove ALL Html tags from our text box. People can send
    ' all sorts of crazy stuff over the web-browser control. including code that can
    ' access your computer if used incorrectly
    txtSend.Text = Replace(txtSend.Text, "<", "&lt;")
    txtSend.Text = Replace(txtSend.Text, ">", "&gt;")

    ' Check to see if our message isn't an action
    If Left(txtSend.Text, 7) <> "/action" And Left(txtSend.Text, 3) <> "/me" Then
        ' Send our message. msg, message to send, message colour, is it bold, is it underlined? ' is it highlighted
        Broadcast "msgø" & txtSend.Text & "ø" & MessageColour & "ø" & MessageBold & "ø" & MessageUnderline    '& "ø" & MessageHColour
        ' Is this message the same as your last one?
        If OldData <> txtSend.Text Then
            ' No? Then give you an NCredits
            ' and remember your last message sent
            OldData = txtSend.Text
            NCredits = NCredits + 1
            
        End If
    ElseIf Left(txtSend.Text, 7) = "/action" Then
        ' Is our message an /action or /me?

        ' Because of the different lengths of the string
        ' (ie. /action = 7, /me = 3), we need to have
        ' different actions for them, namely the mid part

        ' Send it. act, username, message to act out is the syntax
        Broadcast "actø" & Mid(txtSend.Text, 8)
        If OldData <> txtSend.Text Then
            OldData = txtSend.Text
            ' Sending an /action or /me gives you 2!!
            NCredits = NCredits + 2
        End If
    ElseIf Left(txtSend.Text, 3) = "/me" Then
        Broadcast "actø" & Mid(txtSend.Text, 4)
        If OldData <> txtSend.Text Then
            OldData = txtSend.Text
            ' Sending an /action or /me gives you 2 NCredits!!
            NCredits = NCredits + 2
        End If
    End If

    ' Adds the last message to the chat 'history' box
    txtSend.AddItem txtSend.Text, 0
    txtSend.Text = ""
End Sub

Private Sub DoLock()
    On Error Resume Next
    ' When you set your status as AFK or Unavailable,
    ' then everything is locked off except for these
    ' objects. When you set your status as AFK
    ' or unavailable, we don't want you harassing
    ' the people in the room, do we? (Unless you
    ' are an administrator >:D )

    ' BTW, List1 is actually a ListView, not
    ' a generic list box. It is called List1
    ' because it saved changing all the associated code
    Dim Aobject As Object

    For Each Aobject In Me
        Aobject.Enabled = False
    Next

    List1.Enabled = True
    mnuInfo.Enabled = True
    mnuUnAvailable.Enabled = True
    mnuAFK.Enabled = True
    mnuAvailable.Enabled = True
    mnuUE.Enabled = True
    mnuDev.Enabled = True
    mnuUEGhost.Enabled = True
    mnuUEPIP.Enabled = True
    mnuUERedirect.Enabled = True
    mnuUEKick.Enabled = True
    mnuUEAdmin.Enabled = True
    mnuUERemAdmin.Enabled = True
    mnuRawData.Enabled = True
    mnuTray.Enabled = True
    mnuTray_AFK.Enabled = True
    mnuTray_Available.Enabled = True
    mnuTray_Unavailable.Enabled = True
    mnuTray_Show.Enabled = True
    mnuTray_Quit.Enabled = True

End Sub



Private Sub Form_Load()
    On Error Resume Next
    If FileObj.FolderExists(GetTempPath & "User Icons") = False Then MkDir (GetTempPath & "User Icons")
    For I = 1 To ImageList1.ListImages.Count
        SavePicture ImageList1.ListImages(I).Picture, GetTempPath & "User Icons\" & ImageList1.ListImages(I).Key & ".gif"
    Next I


    ' Clear our web-browser window. That way we can insert HTML
    iChat.Navigate "about:blank"
    ' Our folder for incoming custom icons
    'If FileObj.FolderExists(AppPath & "User Icons") = False Then MkDir AppPath & "User Icons"

    ' Set our RoomID (Room Port)
    RoomID = frmMain.sckUDP.RemotePort
    ' Room = RoomID, but I don't know why
    Room = RoomID
    ' Our room name
    RoomName = "Lobby"

    ' Prepare our scripting control for use. This sub is found in
    ' modScripting. It simply lets us use cSubs (The class module)
    ' through our scNotch scripting control
    PrepareScripting

    ' Set our misc form's file browser to the User Icons folder
    'frmMisc.File1.Path = AppPath & "User Icons"
    ' Unload our image list from our Listbox
    'Set frmMain.List1.SmallIcons = Nothing
    ' Add everything in the "User Icons" folder into our imagelist
    'For i = 0 To frmMisc.File1.ListCount - 1
    'frmMain.ImageList1.ListImages.Add 1, frmMisc.File1.List(i), stdole.LoadPicture(AppPath & "User Icons\" & frmMisc.File1.List(i), 64, 64, Default)
    'Next i
    ' Re-set our imagelist to our list1
    'Set frmMain.List1.SmallIcons = frmMain.ImageList1

    ' Tray is actually our class module, clsTray, set in Sub Main
    ' Initialize Syntax: Hwnd to create icon for, Icon
    ' to use, Default tooltip
    Tray.Initialize Me.hwnd, Me.Icon, "NChat - Connecting..."
    Tray.ShowIcon

    ' CaptionPrefix, is what the frmMain.caption starts
    ' as. Eg, if CaptionPrefix = "Hi-", then frmMain's
    ' caption would look like this: "Hi-Grayda has entered
    ' the room". If EndMessage is "!!" then it would
    ' appear like this: "Hi-Grayda has entered the room!!"
    CaptionPrefix = "Welcome to NChat Alpha!! - "
    'MkDir App.Path & "\Recieved Files"
    ' End window is the window's title suffix
    EndWindow = "!!"
    ' Stops other programs from stealing your port :)
    'sckUDP.Bind sckUDP.LocalPort, sckUDP.LocalIP
    Load frmLog

    ' Command Line stuff. use -r#### to load a new room
    ' so you can administer 2+ rooms at the same time
    If Left(Command$, 2) = "-r" Then
        ' NewRoom closes the current connection, changes the port
        ' and then reconnects, in a different "room".
        ' The second part of the command is the name of the room
        ' and the third part is whether or not to announce the
        ' room change (That is optional)
        NewRoom Val(Mid(Command$, 3)), "Startup room"
    Else
        ' If you aren't connecting to another room, then
        ' Tell the user they have 2 open
        If App.PrevInstance = True Then
            MsgBox "It seems you already have NChat open! Having 2 or more NChats open on one computer can cause conflicts, doubled up messages and other nasty stuff. NChat will now close...", vbCritical, "Hey, NChat is already open!"
        '    End
        End If

    End If

    NewRoom "4442", "Lobby", True
    MkDir AppPath & "Profiles"
    ' Set up our file Receiver
    'Receiver.BinaryReceiver1.Listen

    ' Got no profiles in your NChat Alpha Profiles folder? Copy the default
    ' one from our .RES file
    If FileObj.FileExists(AppPath & "Profiles\Default\Index.htm") = False Then
        MkDir AppPath & "Profiles\Default"
        CopyFromRes "101", "PROFILE", "\index.htm", "Profiles\Default"
        Profile = AppPath & "Profiles\Default\Index.htm"
    End If

    ' Loads settings from your settings file. If this is called at any other time,
    ' then the settings set out in frmWelcome (on 1st run) are lost or not changed
    'Do Until UserName <> ""
    LoadSettings
    'Loop
    ' If no icon is loaded in LoadSettings, and by some chance is not set, then
    ' set the icon here.
    If MyIcon = "" Then MyIcon = "Default"
    IP = sckUDP.LocalIP

    ' No profile loaded? Load the default one.
    If Trim(Profile) = "" Then
        Profile = AppPath & "Profiles\Default\index.htm"
    End If

    ' Clears the chat text and displays some headings
    mnuClear_Click
    ' And adds your name to the top of the list
    List1.ListItems.Add 1, frmMain.sckUDP.LocalIP, UserName, , ImageList1.ListImages.Item(FindIcon(MyIcon)).Key

    ' Tell the room you have connected.
    ' Your own 'con' data is not parsed, but instead "Text"ed, because
    ' on some networks, the 1st packet is NOT recieved. I forget the real reason :)
    Broadcast "conø" & MyIcon & "ø" & sckUDP.LocalIP
    DoEvents
    DoEvents
    Text "+username+ has entered the conversation..." & vbCrLf, "con", True, , , 2, "Center"

    ' Status changes our frmMain caption,
    Status CaptionPrefix & UserName & " has entered the conversation..." & EndWindow

    ' For more info on how to tell how fast a program is
    ' loaded, check modMisc, and look at the GetTickCount
    ' Public Declare.
    Log "NChat sucessfully loaded in: " & (GetTickCount - OldTickCount) / 1000 & " milliseconds" & vbCrLf & vbCrLf, vbBlack, True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Code I swiped from somewhere (I think PSC :D)
' That lets you right click on a tray icon

' I don't understand it too well, but it works! :)
    Dim msgCallBackMessage As Long
    msgCallBackMessage = x / Screen.TwipsPerPixelX
    'Me.Caption = msgCallBackMessage
    Const WM_RBUTTONUP = &H205
    Const WM_LBUTTONDBLCLK = &H203

    Select Case msgCallBackMessage
        ' What to do when our tray icon is clicked.
    Case WM_RBUTTONUP
        ' Right mouse button up? Show the menu letting you change your status etc.
        Me.PopupMenu mnuTray
    Case WM_LBUTTONDBLCLK
        ' But if it's a double left click, then toggle the visibility of frmMain
        If mnuTray_Show.Visible = False Then
            mnuHide_Click
        Else
            mnuTray_Show_Click
        End If
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    ' Just some resize stuff
    iChat.Move 150, 150, Me.Width - 400 - List1.Width - 150, Me.Height - 1400 - SB1.Height
    txtSend.Move 150, iChat.Height + 250, iChat.Width
    cmdSend.Move txtSend.Width + 400, txtSend.Top
    cmdAction.Move cmdSend.Left + cmdSend.Width + 50, cmdSend.Top
    ListView1.Move List1.Left, List1.Top, List1.Width, List1.Height
    lstIgnore.Move List1.Left, List1.Top, List1.Width
    'List1.Move iChat.Width + 300, iChat.Top, , iChat.Height
    List1.Move iChat.Width + 300, iChat.Top, List1.Width, iChat.Height
End Sub

Private Sub Form_Terminate()
    Form_Unload (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
' This re-written sub saves our username, icon, NCredits etc. into a .NDAT file
' for later use. This code is about 50% shorter and about 75% faster than the other
' code. It's sorta easily updatable too, if you like to code more crap in :)

'On Error GoTo ErrorH
    Dim TheFile As String

    ' Removes our icon from the tray. No parameters required
    Tray.HideIcon

    ' Tells everyone you have disconnected
    Broadcast "dis"
    DoEvents
    DoEvents
    If IsInternet = True Then sckUDP.Close
    Close #1
    ' Open the file. apppath simply removes or adds a "\" or "/" depending on the location
    Open AppPath & GetUserName & ".ndat" For Output As #1

    ' Writes a "checksum" of our file. This ensures that your file is the latest version
    ' and can utilise the latest NChat features
    TheFile = TheFile & FileCheck & "/©/"

    ' Prepares the second batch of stuff. The quickest and smallest way to do this :)

    ' In the second batch:

    ' Your username
    ' The number of Icons you have purchased from the shop
    TheFile = TheFile & Replace(UserName, "[RA]", "") & "/©/" & TotalIcons & "/©/"

    ' Our admin status. AdminAhoy and TharSheBlows was inspired from the Simpsons. :)
    ' This is done seperatly because of the IF statements required
    If mnuDev.Visible = True Then
        TheFile = TheFile & "AdminAhoy" & "/©/"
    Else
        TheFile = TheFile & "TharSheBlows" & "/©/"
    End If

    ' Third batch of data to be written:

    ' How many NCredits you have
    ' Your Username Check
    ' Your icon
    ' Swearing on or off?
    ' The last loaded NChat Profile
    ' Are you a true Admin?
    ' Show tip of the day on next startup?
    ' Is your username (in iChat) bold?
    ' "                           " Underlined?
    ' "                           " Coloured?
    ' Your message start (eg. || )
    ' Your message end (eg. || )
    TheFile = TheFile & NCredits & "/©/" & GetUserName & "/©/" & Trim(MyIcon) & "/©/" & Swearing & "/©/" & Profile & "/©/" & TrueAdmin & "/©/" & DontShowTip & "/©/" & MessageBold & "/©/" & MessageUnderline & "/©/" & MessageColour & "/©/" & StartMSG & "/©/" & EndMSG & "/©/" & Smiley & "/©/"

    ' Are you banned from NChat?
    If Ban = True Then TheFile = TheFile & "BANNED"

    ' Add the final delimiter, so we can fill the next ~255 bytes of the file with junk.
    ' Yet another security message
    TheFile = TheFile & "/©/"

    ' Write it all to the file, and close it
    Print #1, Encode(TheFile, GetUserName & "€€NCHAT_SOLID¶¶" & GetUserName)

    Close #1

    ' Fade the window out. Thanks to allapi.net for
    ' this code :). This also lets you know if
    ' your settings have been saved
    AnimateWindow Me.hwnd, 200, AW_HIDE Or AW_BLEND
    End
ErrorH:

    MsgBox "There was an error writing your NChat settings file: " & Err.Description & ". Please make sure you aren't using NChat on a CD or other non-writable media. NChat will close, but your settings won't be restored next time you use NChat. Sorry", vbCritical, "Error Writing File"
    End

End Sub


Private Sub iChat_TitleChange(ByVal Text As String)
' This lets our profile call basic NChat commands from the HTML by changing the
' Title of the window.
    Select Case LCase(Text)

    Case "Options"
        frmOptions.Show

    Case "Store"
        frmStore.Show

    Case "Join"
        mnuJoinCustom_Click

    Case "LoadPro"
        mnuLoadProfile_Click

    Case "ClearText"
        mnuClear_Click

    Case "About"
        frmAbout.Show

    Case "JoinNet"
        frmServerConnect.Show

    End Select
End Sub

Private Sub List1_DblClick()
' When you double click a name in the list of users, it opens up a new chat box
    If List1.SelectedItem.Text > "" And List1.SelectedItem.Text <> UserName And List1.ListItems.Count > 0 Then ul1_Click
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    ' If it's not a left click, or it's blank, then ignore it
    If List1.SelectedItem.Text = "" Or Button <> 2 Or List1.ListItems.Count = 0 Then Exit Sub    'Or List1.SelectedItem.Text = Username Then Exit Sub
    If List1.SelectedItem.Text = UserName Then    ' And mnuDev.Visible = False Then
        ' It's a menu, so it pops up under the cursor
        ' This menu has your status, ignore, change
        ' username and icon stuff.
        PopupMenu mnuInfo

        ' Without this, when you right click on your username,
        ' The first menu pops up, then when you click away, the
        ' second one pops up, posing a potential security risk
        Exit Sub
    End If

    ' AFK is Away From Keyboard.
    ' If they are afk, then you can't send 'em messages
    ' unless you are an admin

    ' The VERY last icon in the ImageList1, is the AFK icon (keyboard with red circle)
    ' and the second last one is the Unavailable Icon (Just a red circle)
    If mnuDev.Visible = False And List1.SelectedItem.SmallIcon = ImageList1.ListImages.Item(ImageList1.ListImages.Count).Key Then
        Text "This user is AFK, and cannot be contacted" & vbCrLf, "ThatsBad", True
        Exit Sub
    ElseIf mnuDev.Visible = False And List1.SelectedItem.SmallIcon = ImageList1.ListImages.Item(ImageList1.ListImages.Count - 1).Key Then
        Text "This user is unavailabe, and cannot be contacted" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If
    ' Replaces >> with the username, so you know who
    ' you clicked on
    If List1.SelectedItem.Index > -1 And List1.SelectedItem.Text > "" Then
        ul1.Caption = "Send " & List1.SelectedItem.Text & " a Private Message"
        ul2.Caption = "Send " & List1.SelectedItem.Text & " some NCredits"
        ul4.Caption = "Browse " & List1.SelectedItem.Text & "'s Files"
        ' lol wow, check out the next line. That feature was removed in an unreleased v5.0.2!!
        'ul5.Caption = "Add " & List1.SelectedItem.Text & " to friend list"
        mnuUEIgnore.Caption = "Ignore " & List1.SelectedItem.Text

        ' Are you an admin? if you are, then show the 'extra' menu
        If mnuDev.Visible = True Then mnuUE.Visible = True
        PopupMenu mnuUserList
    End If

End Sub

Private Sub ListView1_DblClick()
' When you double click on a private message, ALL of the messages from that ONE user
' are opened, instead of just one. This stops messages from just popping up at
' inconvinient times

    Dim CurrentUser As String
    Dim Temp As Integer

    If ListView1.SelectedItem.Text = "" Then Exit Sub
    CurrentUser = ListView1.SelectedItem.Text

    For I = 1 To ListView1.ListItems.Count
        If I > ListView1.ListItems.Count Then Exit For
        ListView1.ListItems(I).Selected = True

        ' This is our private message. Part 0 is our key header
        Dim PM() As String

        PM = SplitVB5(ListView1.SelectedItem.Key, "…")

        If ListView1.SelectedItem.Text = CurrentUser Then

            ' FindChatWindow lets us find someone's private chat window by returning it as an integer

            Temp = FindChatWindow(CurrentUser)
            If Temp > 0 Then
                CW(Temp).Show
                ' Add text etc.
                Txt2 ListView1.SelectedItem.Text & " ::  " & PM(1) & vbCrLf, act, Int(Temp)
                ListView1.ListItems.Remove I
                I = 1
            Else
                ' FindFreeWindow finds us the first available chat window
                FreeWindow = FindFreeWindow
                If FreeWindow = 0 Then
                    MsgBox "There are no free windows available! Please close some private message windows to free up slots", vbCritical, "No Free Private Message windows available!"
                    Exit Sub
                End If

                CW(FreeWindow).Show
                CW(FreeWindow).Tag = CurrentUser
                ' Add text etc.
                Txt2 ListView1.SelectedItem.Text & " ::  " & PM(1) & vbCrLf, act, Int(FreeWindow)
                ListView1.ListItems.Remove I
                I = 1
            End If
        End If
    Next I
End Sub

Public Sub LoadProfile()
'On Error Resume Next
' The loading of the profile
' Put into one easy to use sub,
' so the code doesn't have to be repeated
    On Error Resume Next

    Open Profile For Input As #1
    HTMLFile = Input(LOF(1), 1)
    Close #1

    EndWindow = ""
    'tsk. so many doevents... sad.
    DoEvents
    DoEvents

    ' Simple really. GetElement(element to get as string, HTML Text as string) simply
    ' looks in the CSS part of the HTML file, for the colour (eg. .MSG { ), then
    ' gets the 'color: #FFAA00;' part from within the .MSG { }, then turns it into a long
    Heading = GetElement("heading", "color", HTMLFile)
    Msg = GetLongRGB(GetElement("msg", "color", HTMLFile))
    dis = GetLongRGB(GetElement("dis", "color", HTMLFile))
    svr = GetLongRGB(GetElement("svr", "color", HTMLFile))
    act = GetLongRGB(GetElement("act", "color", HTMLFile))
    con = GetLongRGB(GetElement("con", "color", HTMLFile))
    ThatsGood = GetLongRGB(GetElement("thatsgood", "color", HTMLFile))
    ThatsBad = GetLongRGB(GetElement("thatsbad", "color", HTMLFile))
    frmMain.BackColor = GetLongRGB(GetElement("windowback", "color", HTMLFile))
    frmMain.txtSend.BackColor = GetLongRGB(GetElement("sendback", "color", HTMLFile))
    frmMain.txtSend.ForeColor = GetLongRGB(GetElement("sendfore", "color", HTMLFile))
    List1.BackColor = frmMain.txtSend.BackColor
    List1.ForeColor = frmMain.txtSend.ForeColor
    ListView1.BackColor = List1.BackColor
    ListView1.ForeColor = List1.ForeColor
    lstIgnore.BackColor = List1.BackColor
    lstIgnore.ForeColor = List1.ForeColor

    DoEvents
End Sub

Public Sub LoadSettings()
'On Error Resume Next
' Our settings to be applied
    Dim Settings() As String
    ' Our encoded, then decoded settings file
    Dim TheFile As String
    Close #1
    ' Open the file for reading (input)
    If FileObj.FileExists(AppPath & GetUserName & ".ndat") = False Then GoTo ShowWelcome
    Open AppPath & GetUserName & ".ndat" For Input As #1
    DoEvents
    DoEvents
    ' Load it into our string, which can be of almost any length
    TheFile = Input$(LOF(1), 1)
    DoEvents
    DoEvents
    ' Decode the file so it's plaintext
    TheFile = Decode(TheFile, GetUserName & "€€NCHAT_SOLID¶¶" & GetUserName)
    Close #1

    ' Split the file up according to our delimiter
    Settings = SplitVB5(TheFile, "/©/")

    ' This doevents ensures all of our settings are actually loaded.

    ' Check our settings file to make sure it's up-to-date
    ' FileCheck is a constant that tells us our settings file version
    If FileCheck <> Settings(0) And Settings(0) > "" Then MsgBox "This settings file is either out of date, or is invalid. NChat can still continue, but may not provide full compatibility. Danger: Low", vbExclamation, "Bad File Check!"

    ' Are you files shared? (Needs to be converted to a boolean to avoid byref errors)
    'ShareMyFiles = CBool(Settings(1))
    '' Our folder that is shared
    'frmFileList.File1.Path = Settings(2)
    'FilePattern = Settings(3)

    UserName = Settings(1)
    ' No username set? Then it's set as your windows username
    If Trim(UserName) = "" Then
        UserName = GetUserName


        ' using windows 95 / 98 or no windows username set? Then
        ' why not use a random number? :D
        Randomize
        If Trim(UserName) = "" Then UserName = "NChat User #" & Int(Rnd * 10000)
        DoEvents
        DoEvents
        ' Uh, if you have no username, then you are a new user
        ' so show the welcome box for them to punch in their info
ShowWelcome:
        frmWelcome.Show
        ' Loop until you have set up NChat
        Do Until frmWelcome.Visible = False
            'frmMain.WindowState = 1
            DoEvents
            DoEvents
        Loop
        frmTip.Show
    End If
    ' This is how many icons you can pick from the list on frmOptions. YOu can buy more
    TotalIcons = Settings(2)
    ' No icons to pick? Set them up with the default icons
    If TotalIcons < 19 Then TotalIcons = 19

    If Settings(3) = "TharSheBlows" Then
        mnuDev.Visible = False
    Else
        mnuDev.Visible = True
    End If

    NCredits = Settings(4)

    ' Username check stops cheating. NChat takes this setting (For example 'Billy') and
    ' compares it to your current windows username. If they are different, then you may
    ' have tried to dishonestly earn NCredits.
    Usernamecheck = Settings(5)
    If Usernamecheck <> GetUserName And Usernamecheck <> "" Then
        mnuDev.Visible = False
        TrueAdmin = False
        frmCheating.Show
        ' Wait until they have read the box
        Do Until frmCheating.Visible = False
            DoEvents
            DoEvents
        Loop

        ' humiliation!!
        Broadcast ("svrø+username+ tried to earn NCredits dishonestly...")
        Exit Sub
        'End ' Uncomment this line to kick them instead
    End If

    If Trim(Settings(6)) <> "" Then
        MyIcon = Settings(6)
    Else
        MyIcon = "Default"
    End If

    ' Is swearing on?
    Swearing = CBool(Settings(7))

    ' Your last profile loaded
    Profile = Settings(8)

    ' If you have NO profile, then one is loaded AFTER this sub, for people who are
    ' using NChat for the FIRST time.

    If Profile > "" Then mnuClear_Click
    ' True admins cannot be kicked from the NChat Chat room, or have their rights removed
    TrueAdmin = CBool(Settings(9))

    If Settings(10) = "True" Then
        DontShowTip = True
    Else
        frmTip.Show
        OnTop frmTip.hwnd

    End If

    MessageBold = CBool(Settings(11))
    MessageUnderline = CBool(Settings(12))
    MessageColour = Settings(13)

    StartMSG = Settings(14)
    EndMSG = Settings(15)

    If Trim(StartMSG) = "" Then StartMSG = "||"
    If Trim(EndMSG) = "" Then EndMSG = "||"
    Smiley = CBool(Settings(16))
    Banned = Settings(17)
    If Banned = "BANNED" Then
        MsgBox "You have been banned from NChat. NChat will now close", vbCritical, "Banned from NChat"
        End
    End If
    frmMain.Show
    frmSplash.Show

    Close #1
End Sub

Private Sub lstIgnore_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ' Popup the ignore menu, where you can delete ignored people on your list
    If Button = 2 Then PopupMenu mnuPopup_Ignore

End Sub

Private Sub mnuAbout_Click()
' New in this build: New About Box. Used to use API to show
' windows about box, but scrapped that idea, coz I needed to thank
' some people etc.
    frmAbout.Show
End Sub

Private Sub mnuAdminWindow_Click()
' Show the sData log window. Lets you view all incoming NChat data
    frmLog.Show
End Sub

Private Sub mnuAFK_Click()
' Locks NChat down so you can't be afk and send messages
    On Error Resume Next
    Broadcast "chiøAFKø" & IP

    DoLock
End Sub

Private Sub mnuAutoBot_Click()
' Starts / Stops / Shows the NChat Automatic Robot, Called the AutoBot, or
' Notch, as is he / she / it is known
    frmAutobot.Show
End Sub

Private Sub mnuAvailable_Click()
' Sets your status to available
    On Error Resume Next
    ' Chi Lets everyone know that your Icon has changed.
    ' change icon syntax: chi <delimiter>, username <delimiter> new icon
    Broadcast "chiø" & MyIcon & "ø" & IP
    ' UnLocks NChat
    Dim Obj As Object
    For Each Obj In frmMain
        Obj.Enabled = True
    Next
    txtSend.Text = ""
End Sub

Private Sub mnuAwayMSG_Click()
' Away message is like leaving a note for any potential
' Private Messengers. It is also like an answering machine
    If AwayMessage = False Then

        AwayMSG = InputBox("Enter an Away Message. If someone tries to contact you in a private chat, this message will be displayed", "Away Message", AwayMSG)
        mnuAwayMSG.Checked = True
        AwayMessage = True
    Else
        mnuAwayMSG.Checked = False
        AwayMessage = False
    End If
End Sub

Private Sub mnuBalance_Click()
' < 0 NCredits, red text > 0 NCredits green text
    If NCredits <= 0 Then
        Text "You have +ncredits+ NCredits remaining" & vbCrLf, "ThatsBad", True
    Else
        Text "You have +ncredits+ NCredits remaining" & vbCrLf, "ThatsGood", True
    End If
End Sub

Private Sub mnuBL_Click()
' Broadcast Mode lets you chat to EVERYONE on the network. Loopback lets
' NChat connect to itself, making it useful for testing on one computer,
' if you don't have a valid network
    If MsgBox("WARNING: BY SWITCHING TO LOOPBACK MODE, NCHAT CAN CONNECT TO ITSELF, AND NOT WITH OTHER NCHATS ON OTHER COMPUTERS. ONLY DO THIS IF YOU KNOW WHAT YOU ARE DOING. IF YOU ARE UNSURE, CLICK CANCEL", vbCritical + vbYesNo, "WARNING: READ CAREFULLY") = vbNo Then Exit Sub
    If Loopback = False Then
        Broadcast "dis"
        Address = "127.0.0.1"
        Loopback = True
        MsgBox "Now in loopback mode. To change back, click button again", vbInformation, "Loopback"
        Broadcast "conø" & MyIcon & "ø" & IP
        'frmMain.SB1.Panels(1).Text = "NChat - Online!!"
        'frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture
    Else
        Broadcast "dis"
        Address = "255.255.255.255"
        Loopback = False
        MsgBox "Now in broadcast mode. To change back, click button again", vbInformation, "Loopback"
        Broadcast "conø" & MyIcon & "ø" & IP
        'frmMain.SB1.Panels(1).Text = "NChat - Online!!"
        'frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture
    End If
End Sub

Private Sub mnuChangeNCredits_Click()
' Change your NCredits, and stops any overflow errors
    On Error Resume Next
    NCredits = InputBox("How many NCredits would you like?", "Change NCredits", NCredits)
End Sub

Private Sub mnuChangeRoomInfo_Click()
' Is as it says. Changes your room info if you make a spelling mistake
    Tmp = InputBox("Please enter the new name for your room", "New Room Name", RoomName)
    If Trim(Tmp) = "" Then
        Text "Room name was blank. Not Changed" & vbCrLf, "ThatsBad", True
        Exit Sub
    ElseIf Trim(Tmp) = RoomName Then
        Text "Room name is the same as your old one. Not Changed" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If
    Text "This room (" & RoomName & ") has been changed to " & Tmp & "." & vbCrLf, "ThatsGood", True
    Broadcast "svrøThis room's name is now set as: " & Tmp & " (Was " & RoomName & ")"
    RoomName = Tmp

End Sub

Private Sub mnuClear_Click()
'Dim HTMLFile As String

'Open Profile For Input As #1
'HTMLFile = Input(LOF(1), 1)
'Close #1
' Replaces ALL innerHTML with our HTML File. When other HTML is inserted,
' The new HTML is appended to the end
    LoadProfile
    'frmMain.iChat.Document.body.innerHTML = HTMLFile
DoEvents
    iChat.Navigate Profile
DoEvents
End Sub

Private Sub mnuConnect_Click()
    frmServerConnect.Show

End Sub

Private Sub mnuCreateRoom_Click()
    frmCRI.Show
End Sub

Private Sub mnuDeleteIgnore_Click()
    If lstIgnore.SelectedItem.Text > "" Then lstIgnore.ListItems.Remove (lstIgnore.SelectedItem.Text)
End Sub

Private Sub mnuDisconnect_Click()
    IsInternet = False
    Address = "255.255.255.255"
    sckUDP.Close
    sckRooms.Close
    sckUDP.Protocol = sckUDPProtocol
    sckRooms.Protocol = sckUDPProtocol
    sckUDP.Connect
    sckRooms.Connect
    mnuChatRooms.Visible = True

End Sub

Private Sub mnuDoHeading_Click()
' Large 'message' (Heading)
    Broadcast "heaø" & InputBox("What do you want your heading to say?", Heading) & "ø" & UserName
    ' 'Resets' the text so the next message isn't BIG
    Text "" & vbCrLf, , , , , 2

End Sub



Private Sub mnuDownload_Click()
    frmDownloads.Show

End Sub

Private Sub mnuHide_Click()
    mnuTray_Show.Visible = True
    mnuHide.Visible = False

    Me.Visible = False
    Tray.Box "NChat has been minimized to the tray! To restore it, double click the NChat icon!", "NChat"
End Sub

Private Sub mnuIgnore_Click()
    mnuIgnore2_Click

End Sub

Private Sub mnuIgnore2_Click()
    List1.Visible = False
    lstIgnore.Visible = True
    ListView1.Visible = False
    mnuUL.Checked = False
    mnuPM.Checked = False
    mnuIgnore2.Checked = True
End Sub

Private Sub mnuInsertFake_Click()
' User (Fake or real) connect syntax:
'  con <Delmiter> Username <Delmiter> Icon #
    On Error Resume Next
    Dim Icon As Integer
    Randomize
    FakeUser = InputBox("Enter fake user's name", "Fake user")
    If FakeUser > "" Then
        Broadcast "fakeuø" & FakeUser & "ø" & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255)
        'Tex6t FakeUser & " has entered the conversation!" & vbCrLf, con, True, , , , "Center"
        'List1.ListItems.Add 2, FakeUser, FakeUser, , ImageList1.ListImages.Item(Icon).Key
        NewUser = FakeUser
    End If
End Sub

Private Sub mnuJoinCustom_Click()
    On Error Resume Next
    frmRooms.List1.ImageList = ImageList1

    frmRooms.List1.Nodes.Clear
    frmRooms.List1.Nodes.Add , , "N", "NChat Rooms", frmMain.ImageList1.ListImages.Item(18).Key
    frmRooms.List1.Nodes.Add , , "C", "User-Made Rooms", frmMain.ImageList1.ListImages.Item(18).Key
    frmRooms.List1.Nodes.Item(1).Expanded = True
    frmRooms.List1.Nodes.Item(2).Expanded = True
    RI = frmMain.ImageList1.ListImages.Item(7).Key
    frmRooms.List1.Nodes.Add "N", 4, "R4442", "Lobby", RI
    frmRooms.List1.Nodes.Add "N", 4, "R4443", "Music Chat", RI
    frmRooms.List1.Nodes.Add "N", 4, "R4444", "The Work Room", RI
    frmRooms.List1.Nodes.Add "N", 4, "R4445", "Help for NChat", RI
    frmRooms.List1.Nodes.Add "N", 4, "R4446", "Programmers Chat", RI
    frmRooms.List1.Nodes.Add "N", 4, "R4447", "Room for fighters", RI
    RoomBroadcast "lst"
    frmRooms.Show
End Sub

Private Sub mnuLoadProfile_Click()
'dlgSave.Filter = "NChat Profile (*.pro)|*.pro|HTML File (*.html, *.htm)|*.html;*.htm|All files (*.*)|*.*"
'dlgSave.DialogTitle = "Select NChat profile to load"
'dlgSave.ShowOpen

'If dlgSave.FileName <> "" Then
'Profile = dlgSave
'mnuClear_Click
'If OldINI <> Profile Then
'NCredits = NCredits + 10
'Text "You have recieved 10 NCredits for loading a profile!!" & vbCrLf, ThatsGood, True, , , , "Center"
'End If

'End If

    Temp = BrowseForFolder
    If Temp = "" Then Exit Sub

    If Right(Temp, 1) <> "\" Then Temp = Temp & "\"
    If FileObj.FileExists(Temp & "index.htm") = False Then
        MsgBox "The folder: " & Temp & " is missing a critical file. Index.htm is missing, which is used for text colours etc and let NChat change it's colours for windows and forms. Please check the file exist, and try again", vbCritical, "Missing Critical Profile File!"
        Exit Sub
    End If

    Profile = Temp & "index.htm"
    mnuClear_Click
    If OldINI <> Profile Then
        NCredits = NCredits + 10
        Text "You have recieved 10 NCredits for loading a profile!!" & vbCrLf, "ThatsGood", True, , , , "Center"
    End If
End Sub

Private Sub mnuNewRoom_Click()
' Creates a room. Doesn't have a random # like above
    Tmp = InputBox("Enter the room to create. You can choose the room number", "Custom Room")
    RoomName = InputBox("Enter your room's name. You can choose the new name", "Custom Room")
    If Tmp > "" Then
        CreatedRoom = True

        NewRoom Tmp, RoomName
        
    End If

End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show

End Sub

Private Sub mnuPM_Click()
    List1.Visible = False
    lstIgnore.Visible = False
    ListView1.Visible = True
    mnuUL.Checked = False
    mnuPM.Checked = True
    mnuIgnore2.Checked = False


End Sub

Private Sub mnuQuit_Click()
    Form_Unload (0)
End Sub

Private Sub mnuRAKick_Click()
' Showbox brings up a box with all users in the room in it
' So we can pick a name, and click it, without having to
' copy a username, paste it and click OK
    ShowBox "Kick User", "Kick a user from NChat"
    Broadcast "ksvø" & SelUser
End Sub

Private Sub mnuRAPassword_Click()
    OldPassword = Password
    Password = InputBox("Please enter a new password, or leave it blank to remove the password", "Change password")
    If MsgBox("Are you sure you want to change / remove the password? This will affect new people trying to enter the room", vbCritical + vbYesNo, "Confirm change password?") = vbNo Then
        Password = OldPassword
        Text "Password NOT changed", "ThatsBad", True
    Else
        Text "Password CHANGED! New password: " & Left(Password, 1) & String(Len(Password) - 2, "*") & Right(Password, 1), "ThatsGood", True
    End If

End Sub

Private Sub mnuRAServerMSG_Click()
Dim Tmp As String
' Sends purple (Server) messages
 Tmp = InputBox("Enter text to send as purple", "Send Server Text")
    If Trim(Tmp) > "" Then Broadcast "svrø" & Tmp

End Sub

Private Sub mnuRawData_Click()
    Dim tmp2 As String

    ' Raw data refers to data that doesn't have a prefix such
    ' as hea, msg, cdo ect.
    tmp2 = InputBox("Enter raw data to broadcast. MUST include delimiters for standard messages (Alt+0248 or +d+)...", "Raw Data")
    If tmp2 > "" Then Broadcast tmp2
End Sub

Private Sub mnuRAWelcome_Click()
' Your welcome message

' If no message is set, then say: message for this room SET
' but if the message is being updated, then say so
    If WelcomeMsg = "" Then
        WelcomeMsg = "svr+d+" & InputBox("Enter a welcome message that everyone will see", "Welcome Message", WelcomeMsg)
        Broadcast "svrøWelcome message for this room set!"
    End If
End Sub

Private Sub mnuRealUser_Click()
' isr = Is Real? Checks if a user is real or fake
    Broadcast "isr+d+" & List1.SelectedItem.Text
End Sub

Private Sub mnuRemFake_Click()
' Removes a fake user from the room
    Broadcast "disø" & InputBox("Enter username to disconnect", "Remove Fake User", NewUser)
End Sub

Private Sub mnuSaveChat_Click()
' Saves chat data
' If the selected type is *.txt or *.*
' then save it as plain text.
' If it's not, then save it with RTF tags etc.
    dlgSave.Filter = "HTM Files (*.htm)|*.HTM|All Files (*.*)|*.*"
    dlgSave.ShowSave
    If dlgSave.FileName = "" Then Exit Sub
    Open dlgSave.FileName For Output As #1
    Print #1, iChat.Document.body.innerHTML
    Close #1
    Text "File Saved!" & vbCrLf, svr, True
End Sub

Private Sub mnuSendAll_Click()
    Tmp = InputBox("Enter text to send. This will be sent to EVERYONE using NChat, even people in other rooms", "Send to ALL Rooms")
    If Tmp > "" Then sckRooms.SendData "comø" & Tmp
End Sub

Private Sub mnuSmileys_Click()
' Sometimes all the smileys won't show up, so do it twice
' to ensure it shows up ok
    frmSmileys.Show
    Unload frmSmileys
    frmSmileys.Show
End Sub

Private Sub mnuStore_Click()
    frmStore.Show
End Sub

Private Sub mnuTextHelp_Click()
    If FileObj.FileExists(AppPath & "Help\NChat Help.chm") = True Then
        ShellExecute 0&, "Open", AppPath & "Help\NChat Help.chm", "", vbNullString, 1
    Else
        MsgBox "Cannot find " & AppPath & "Help\NChat Help.chm. The help file cannot be opened. You can download it off the web-site at: http://www.solidinc.tk, under the Downloads section", vbCritical, "Help File Not Found!!"
    End If


End Sub

Private Sub mnuTipofTheDay_Click()
    frmTip.Show

End Sub

Private Sub mnuTray_AFK_Click()
    mnuAFK_Click
End Sub

Private Sub mnuTray_Available_Click()
    mnuAvailable_Click
End Sub

Private Sub mnuTray_Quit_Click()
    Form_Unload (0)
End Sub

Private Sub mnuTray_Show_Click()
    Me.Visible = True
    Me.WindowState = 0
    mnuHide.Visible = True
    mnuTray_Show.Visible = False
End Sub

Private Sub mnuTray_Unavailable_Click()
    mnuUnAvailable_Click
End Sub

Private Sub mnuUEAdmin_Click()
' Add admin (This menu is called from List1 on right click

' UE = User... um... Environment?... Education? Never mind :)
    Broadcast "addø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUECHI_Click()
' Change your icon without going to options etc.
    Tmp = InputBox("Enter a number between 1 and " & TotalIcons & " as your Icon", "Change User Icon", FindIcon(MyIcon))
    If Tmp = FindIcon(MyIcon) Then
        Text "Your Icon is the same as your old one!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If

    ' Cannot enter a -tive number
    If Tmp < 1 Or Tmp = "" Then
        Text "Your New Icon is less than 1!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If

    ' Only admins can access all icons
    If Tmp > ImageList1.ListImages.Count Then
        Text "Your New Icon is more than " & ImageList1.ListImages.Count & "!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If

    ' Cannot enter a # more than the number of icons you own
    If Tmp > TotalIcons And mnuDev.Visible = False Then
        Text "Your New Icon is more than " & TotalIcons & "!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If


    MyIcon = Tmp
    ' Tell the room of your icon change
    Broadcast "chi+d+" & MyIcon & "ø" & IP
End Sub

Private Sub mnuUECHU_Click()
' CHange username
' Admins can have [A] and [RA] on their name
' but not wannabes
    OldUsername = UserName
    Tmp = InputBox("Enter new username. It has to be different to your old one, and cannot be blank", "Change Username", UserName)
    If Tmp = "" Then
        Text "Your New Username is blank!!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If

    If Right(Trim(Tmp), 3) = "[A]" And frmMain.mnuDev.Visible = False Or Right(Trim(Tmp), 4) = "[RA]" And frmMain.mnuDev.Visible = False Then
        MsgBox "Sorry, but only administrators can have [A] or [RA] on the end of their name...", vbExclamation, "Bad Username"
        Exit Sub
    End If


    If Tmp = UserName Then
        Text "Your New Username is the same as your old one!" & vbCrLf, "ThatsBad", True
        Exit Sub
    End If

    UserName = Tmp
    Broadcast "chuø" & OldUsername & "ø" & MyIcon
    DoEvents

    Broadcast "svrø" & OldUsername & " is now known as " & UserName
    Text "Changed your username!" & vbCrLf, Heading, True, , , , "Center"


End Sub

Private Sub mnuUEGhost_Click()
' Ghosts (Message with someone else's username) a user
    Broadcast "fakø" & List1.SelectedItem.Text & "ø" & InputBox("Enter message", "Ghost") & "ø" & UserName
End Sub

Private Sub mnuUEIgnore_Click()
    On Error Resume Next
    ' QuickIgnore(TM) :)
    If MsgBox("Are you sure you want to ignore: " & List1.SelectedItem.Text & "? To unignore them, right click your name and select ignore list.", vbExclamation + vbYesNo, "Ignore User?") = vbYes Then lstIgnore.ListItems.Add , List1.SelectedItem.Text, List1.SelectedItem.Text, , ImageList1.ListImages.Item(ImageList1.ListImages.Count).Key

End Sub

Private Sub mnuUEKick_Click()
' KSV = Kick Server
' Kicks someone with a message from the admin
' (i.e The admin has kicked you etc.)
    Broadcast "ksvø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUEPIP_Click()
' Print detailed info about a user
    Broadcast "pipø" & List1.SelectedItem.Text & "ø" & UserName


End Sub

Private Sub mnuUERedirect_Click()
' Redirect a user.
    RTMP = InputBox("Enter room number to redirect to:", "Redirect")
    RTMP2 = InputBox("Enter room name", "Redirect", "Holiday Destination")
    If RTMP > "" And RTMP2 > "" Then
        Broadcast "redø" & RTMP & "ø" & List1.SelectedItem.Text & "ø" & RTMP2 & "ø" & UserName
        Exit Sub
    End If

End Sub

Private Sub mnuUERemAdmin_Click()
' Kill an admin's rights
    Broadcast "remø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUL_Click()
    List1.Visible = True
    lstIgnore.Visible = False
    ListView1.Visible = False
    mnuUL.Checked = True
    mnuPM.Checked = False
    mnuIgnore2.Checked = False
End Sub

Private Sub mnuUnAvailable_Click()
    On Error Resume Next
    ' Makes sets your icon as 'unavailable'.
    ' Users cannot contact you while you are away
    Broadcast "chiøUnavailableø" & IP

    ' Locks NChat down so you can't be unavailable and send messages
    AwayMessage = True
    DoLock
End Sub

Private Sub mnuWmsg_Click()
' The welcome message is displayed on a user connect
' it doesn't have to be a message, but can be anything
' even a user limiter (i.e ksvø+newuser+)

' This is the same as DoStuff on entry, but
' only one command is sent
    If WelcomeMsg = "" Then
        WelcomeMsg = InputBox("Please enter new AutoWelcome Message", "WelcomeMsg", WelcomeMsg)
        Broadcast ("svrøWelcome message set!!")
    Else
        WelcomeMsg = InputBox("Please enter updated AutoWelcome Message", "WelcomeMsg", WelcomeMsg)
        Broadcast ("svrøWelcome message updated!!")
    End If
End Sub

Private Sub sckRooms_DataArrival(ByVal bytesTotal As Long)
Dim rData As String
sckRooms.GetData rData
ParseRoomData (rData)
End Sub

Private Sub sckUDP_Close()
    Text "You have been disconnected from the server. This could be because the server has crashed, or the administrator has kicked you", vbBlack, True
End Sub

Private Sub sckUDP_Connect()
    If sckUDP.Protocol = sckTCPProtocol Then Broadcast "conø" & MyIcon & "ø" & sckUDP.LocalIP

End Sub

Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)
sckUDP.GetData sData
ParseData sData

End Sub

Public Sub MenuStuff(Menu As Integer)

Select Case Menu

' Close NChat
Case 0
Form_Unload (0)

Case 1
mnuDisconnect_Click

Case 2

End Select
End Sub

Private Sub sckUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If IsInternet = True Then MsgBox "There was an error with your NChat. Please report the following details to Grayda (www.solidinc.tk)" & vbCrLf & vbCrLf & Number & vbCrLf & Description, vbCritical, "NChat Error"

End Sub

Private Sub Timer1_Timer()
' Sends your online presence to everyone. THere
' are better ways to do this, but I can't be
' bothered :P
    Broadcast "usrø" & MyIcon & "ø" & sckUDP.LocalIP & "ø" & CreatedRoom
    If CreatedRoom = True Then
        RoomTime = RoomTime + 1
        If Password > "" Then
            RoomBroadcast "roomø" & "R" & sckUDP.LocalPort & "/ù/" & "ø" & RoomName & "ø" & MyIcon & "ø"
        Else
            RoomBroadcast "roomø" & "R" & sckUDP.LocalPort & "/ù/" & Encode(Password, "FireStormInc") & "ø" & RoomName & "ø" & MyIcon & "ø"
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    On Error GoTo Timer2_Timer_Error

    On Error Resume Next
    'MessageColour = Msg
    If ListView1.ListItems.Count = 0 Then

        Tray.IconHandle = Me.Icon
    Else
        Tray.IconHandle = ImageList1.ListImages.Item(ImageList1.ListImages.Count - 2).Picture
    End If

    ' The room you are in and how long you have been on NChat for
    NChatTime = NChatTime + 1
    SB1.Panels(2).Text = "You are currently in room: " & RoomName & " (#" & RoomID & ")"

    If CreatedRoom = True Then
        hr5.Visible = True
        mnuChangeRoomInfo.Visible = True
        mnuRoomAdmin.Visible = True
        mnuRAKick.Visible = True
        mnuRAPassword.Visible = True
        mnuRAServerMSG.Visible = True
        mnuRAWelcome.Visible = True
        hr12.Visible = True
        
    Else
        hr5.Visible = False
        mnuChangeRoomInfo.Visible = False
        mnuRoomAdmin.Visible = False
        mnuRAKick.Visible = False
        mnuRAPassword.Visible = False
        mnuRAServerMSG.Visible = T
        mnuRAWelcome.Visible = False
        hr12.Visible = False
    End If


    'If ShareMyFiles = False Then
    'NOFS = 0
    'Else
    'NOFS = frmFileList.File1.ListCount
    'End If

    'List1.ListItems.Item(1).SubItems(1) = NOFS
    'If FilePattern > "" Then frmFileList.File1.Pattern = FilePattern

    On Error GoTo 0
    Exit Sub

Timer2_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer2_Timer of Form frmMain"

End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSend_Click

End Sub

Private Sub ul1_Click()
' Send a private message from the user list
    On Error Resume Next


    I = FindFreeWindow

    If I > 0 Then
        CW(I).WindowState = 0
        CW(I).Show
        CW(I).Tag = List1.SelectedItem.Text
        CW(I).Picture1.Tag = List1.SelectedItem.Key
    End If


End Sub

Private Sub ul2_Click()
    On Error Resume Next
    ' Sends some NCredits to the person you clicked on
    ' on the user list
    ToWho = List1.SelectedItem.Text
    HowMany = InputBox("How many NCredits do you want to give?", "Give NCredits")
    If NCredits > HowMany And HowMany > 0 Then
        Broadcast ("sndø" & ToWho & "ø" & HowMany & "ø" & UserName)
        NCredits = NCredits - HowMany
    ElseIf HowMany <> "" Then
        Text "You do not have enough NCredits to give away!!" & vbCrLf, "ThatsBad", True
    ElseIf HowMany = "" Then
        Exit Sub
    End If

End Sub






