VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to NChat Alpha!"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   354.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   477
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   7095
      Index           =   4
      Left            =   2520
      ScaleHeight     =   7095
      ScaleWidth      =   7095
      TabIndex        =   41
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command21 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   5280
         TabIndex        =   43
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "<< Back"
         Height          =   495
         Left            =   3720
         TabIndex        =   42
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWelcome.frx":1B7A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   360
         TabIndex        =   45
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ALL DONE!"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   44
         Top             =   120
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   3
      Left            =   2520
      ScaleHeight     =   7095
      ScaleWidth      =   7095
      TabIndex        =   25
      Top             =   0
      Width           =   7095
      Begin MSComDlg.CommonDialog dlgPro 
         Left            =   3120
         Top             =   6480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select a custom Profile to load OR cancel to quit"
         Filter          =   "NChat Profile (*.pro)|*.pro|All files (*.*)|*.*"
         Flags           =   4
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Browse"
         Height          =   375
         Left            =   5280
         TabIndex        =   40
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   39
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton Command9 
         Caption         =   "?"
         Height          =   375
         Left            =   6480
         TabIndex        =   38
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "?"
         Height          =   375
         Left            =   6480
         TabIndex        =   37
         Top             =   1800
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "<< Back"
         Height          =   495
         Left            =   3720
         TabIndex        =   27
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   5280
         TabIndex        =   26
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Or load your own:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a profile:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NCHAT PROFILE LOADING"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   29
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "On this page, you can load a profile (colour scheme) for NChat's main screen (The GUI). "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   6375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Index           =   2
      Left            =   2520
      ScaleHeight     =   7095
      ScaleWidth      =   7095
      TabIndex        =   14
      Top             =   0
      Width           =   7095
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Swearing off?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "?"
         Height          =   375
         Left            =   6360
         TabIndex        =   32
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Smileys off?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "?"
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Top             =   3240
         Width           =   375
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   2160
         TabIndex        =   24
         Top             =   2280
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "Please select an Icon to use"
      End
      Begin VB.CommandButton Command7 
         Caption         =   "?"
         Height          =   375
         Left            =   6360
         TabIndex        =   23
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "?"
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   1800
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   5280
         TabIndex        =   16
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<< Back"
         Height          =   495
         Left            =   3720
         TabIndex        =   15
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Username:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This page will let you set up your user information, such as your NChat username, your icon for use in NChat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NCHAT USER INFORMATION"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   17
         Top             =   120
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   1
      Left            =   2520
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   9
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "<< Back"
         Height          =   495
         Left            =   3720
         TabIndex        =   13
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmWelcome.frx":1C4B
         Top             =   840
         Width           =   6615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   5280
         TabIndex        =   10
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NCHAT QUICK NOTES"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1800
         TabIndex        =   11
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   0
      Left            =   2520
      ScaleHeight     =   7095
      ScaleWidth      =   7095
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWelcome.frx":22C7
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   6615
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWelcome.frx":235D
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME TO NCHAT ALPHA !"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Image Image1 
      Height          =   2910
      Left            =   45
      Picture         =   "frmWelcome.frx":2474
      Top             =   4125
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FILE SHARING SETUP"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING PROFILES"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER INFORMATION"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "QUICK NOTES"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO NCHAT"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' What step of the wizard are we up to?
Dim TheBox As Integer

Private Sub Command1_Click()
    NextPage
End Sub

Private Sub PrevPage()
    Dim PicBox As Object
    For Each PicBox In Me
        If TypeOf PicBox Is PictureBox Then PicBox.Visible = False
    Next

    For I = 0 To 4
        Label1(I).ForeColor = vbBlack
    Next I

    Label1(TheBox - 1).ForeColor = &H808080

    TheBox = TheBox - 1
    Picture1(TheBox).Visible = True
End Sub


Private Sub NextPage()
    Dim PicBox As Object

    For Each PicBox In Me
        If TypeOf PicBox Is PictureBox Then PicBox.Visible = False
    Next

    For I = 0 To 4
        Label1(I).ForeColor = vbBlack
    Next I
    Label1(TheBox + 1).ForeColor = &H808080


    TheBox = TheBox + 1
    Picture1(TheBox).Visible = True
End Sub

Private Sub Command10_Click()
    MsgBox "If this box is TICKED, then smiley text (often called emoticons) WONT be replaced by the picture (for example. :) will bring up the smiley face). Merely a cosmetic thing, and speeds things up on REALLY old computers", vbQuestion, "Smileys off?"

End Sub

Private Sub Command11_Click()
    MsgBox "If this box is TICKED, then some swearing that comes through NChat will be blocked with asterisks (*). This doesn't stop varients from coming in (such as replacing I with 1, o with 0 etc", vbQuestion, "Swearing off?"

End Sub

Private Sub Command12_Click()
    dlgPro.ShowOpen
    If dlgPro.FileName <> "" Then Text3.Text = dlgPro.FileName

End Sub


Private Sub Command14_Click()
    NextPage
End Sub

Private Sub Command15_Click()
    PrevPage
End Sub

Private Sub Command2_Click()
    NextPage
End Sub

Private Sub Command20_Click()
    PrevPage

End Sub

Private Sub Command21_Click()
    On Error Resume Next
    MsgBox "All Done! NChat will now load!", vbInformation, "Done!"
    If Text2.Text > "" Then
        UserName = Text2.Text
    Else
        UserName = GetUserName
    End If

    If ImageCombo1.SelectedItem.Text > "" Then
        MyIcon = ImageCombo1.SelectedItem.Text
    Else
        MyIcon = "Default"
    End If

    If Check4.Value = True Then
        Swearing = False
    Else
        Swearing = True
    End If

    If Check3.Value = True Then
        Smiley = False
    Else
        Smiley = True
    End If

    If Text3.Text > "" Then
        Profile = Text3.Text & "\index.htm"
    ElseIf Trim(Text3.Text) = "" And Combo1.ListIndex > -1 Then
        Profile = Combo1.List(Combo1.ListIndex) & "\index.htm"
    ElseIf Trim(Text3.Text) = "" And Combo1.ListIndex = -1 Then
    End If

    frmMain.LoadProfile
    Unload Me
    frmMain.Show

End Sub

Private Sub Command3_Click()
    PrevPage
End Sub

Private Sub Command4_Click()
    PrevPage
End Sub

Private Sub Command5_Click()
    NextPage
End Sub

Private Sub Command6_Click()
    MsgBox "This box is for your username. Your username is how others can tell who you are. It can be ANYTHING you want, under 25 letters, and MUST NOT contain [A], (A}, [RA], or (RA). Also, try not to make your username offensive to others in the room", vbQuestion, "Your Username"

End Sub

Private Sub Command7_Click()
    MsgBox "When you select an icon, it will appear next to your username in the list of users. This helps people indentify you even faster by your icon. At the moment, you can't use your own Icon, but this is coming soon!", vbQuestion, "Your Icon"
End Sub

Private Sub Command8_Click()
    MsgBox "NChat comes with 14 default profiles. This box lets you chose one of the default ones", vbQuestion, "Select a profile"

End Sub

Private Sub Command9_Click()
    MsgBox "Here, you can select your own profile to use (if you made one already using the editor). Please note that if you specify 2 profiles (in the drop-down box AND in the text box, then the TEXT BOX ONE WILL BE LOADED, NOT the drop-down one", vbQuestion, "Load your own profile"

End Sub

Private Sub Form_Load()
    TheBox = -1
    NextPage
    
    For Each PicBox In Me
        If TypeOf PicBox Is PictureBox Then PicBox.Picture = Picture1(0).Picture
    Next


    ' Load all the user icons into our image combo box. ONLY LOAD 19
    ' because they have to buy the rest later :D
    ImageCombo1.ImageList = frmMain.ImageList1
    For I = 1 To 19
        ImageCombo1.ComboItems.Add , frmMain.ImageList1.ListImages(I).Key, frmMain.ImageList1.ListImages(I).Key, frmMain.ImageList1.ListImages(I).Key
    Next I
    ' Then load all available profiles into the combo box
    frmMisc.Dir1.Path = AppPath & "profiles"
    For I = 1 To frmMisc.Dir1.ListCount - 1
        Combo1.AddItem frmMisc.Dir1.List(I), 0
    Next I
    Text2.Text = GetUserName

End Sub

