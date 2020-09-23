VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New NChat Profile"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   6795
   Icon            =   "frmNewProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   30
      TabIndex        =   3
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Profile Colours"
      TabPicture(0)   =   "frmNewProfile.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "a1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "aHeading"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "a1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "aMessages"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label17"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label21"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "a1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "a1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "a1(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "a1(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "a1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "a1(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "a1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "a1(8)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "a1(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "a1(12)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dlgColour"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Profile Information"
      TabPicture(1)   =   "frmNewProfile.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(4)=   "Label23"
      Tab(1).Control(5)=   "Text5"
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(7)=   "Command6"
      Tab(1).Control(8)=   "Command4"
      Tab(1).Control(9)=   "List1"
      Tab(1).Control(10)=   "Text4"
      Tab(1).Control(11)=   "Text3"
      Tab(1).Control(12)=   "Text2"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "frmNewProfile.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtChat"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   43
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         TabIndex        =   11
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         TabIndex        =   10
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   4935
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1200
         ItemData        =   "frmNewProfile.frx":2D4E
         Left            =   -73320
         List            =   "frmNewProfile.frx":2D50
         TabIndex        =   8
         Top             =   1920
         Width           =   4935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -71040
         TabIndex        =   7
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear all"
         Height          =   375
         Left            =   -69720
         TabIndex        =   6
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   -72360
         TabIndex        =   5
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         TabIndex        =   4
         Top             =   840
         Width           =   4935
      End
      Begin MSComDlg.CommonDialog dlgColour 
         Left            =   120
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Color           =   33023
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmNewProfile.frx":2D52
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   42
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   41
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   40
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   39
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   38
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   37
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   36
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   34
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   33
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Send fore colour:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Chat back colour:"
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Send back colour:"
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Background Picture:"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Background Colour:"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Error / Bad news:"
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Special / Good:"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "User actions:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Server Messages:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "User Connect:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "User Disconnect:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label aMessages 
         Caption         =   "User Messages:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label aHeading 
         Caption         =   "Headings:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label a1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   18
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "Window Title:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Profile Author:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Profile Description:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Profile Welcome Message (50 Lines)"
         Height          =   495
         Left            =   -74880
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Window End:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmNewProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub A1_Click(Index As Integer)
On Error GoTo NoCol
' When you click on an icon, then
' the index is worked out and
' Changes that colour only, not all of them

dlgColour.ShowColor

a1(Index).BackColor = dlgColour.color
NoCol:
Exit Sub
End Sub



Private Sub Command1_Click()
On Error Resume Next
' Opens (or creates) an INI file, and prints
' your colours for later use
frmPic = Text1.Text

dlgColour.Filter = "INI Files (*.pro)|*.pro"
dlgColour.DialogTitle = "Select Location to save..."

dlgColour.ShowSave
If dlgColour.FileName = "" Then Exit Sub
Open dlgColour.FileName For Output As #1
Print #1, "[Theme]"

Print #1, "Heading=" & CLng(a1(0).BackColor)
Print #1, "Message=" & CLng(a1(1).BackColor)
Print #1, "Disconnect=" & CLng(a1(2).BackColor)
Print #1, "Connect=" & CLng(a1(3).BackColor)
Print #1, "Server=" & CLng(a1(4).BackColor)
Print #1, "Action=" & CLng(a1(5).BackColor)
Print #1, "Good=" & CLng(a1(8).BackColor)
Print #1, "Error=" & CLng(a1(9).BackColor)
Print #1, "Background=" & CLng(a1(10).BackColor)
Print #1, "BackgroundPic=" & Text1.Text
Print #1, "SendBackColor=" & CLng(a1(11).BackColor)
Print #1, "ChatBackColor=" & CLng(a1(12).BackColor)
Print #1, "SendForeColor=" & CLng(a1(6).BackColor)
Print #1, "Title=" & Text2.Text
For i = 0 To 50
If List1.List(i) > "" Then Print #1, "WelcomeMSG" & i & "=" & List1.List(i)
Next i
Print #1, "Author=" & Text3.Text
Print #1, "Description=" & Text4.Text
Print #1, "EndWindow=" & Text5.Text
Close #1

' Sets the path of the ini, then
' Loads your profile so you can see
' if there are any problems
IniFile(1) = dlgColour.FileName
frmMain.LoadProfile

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Preview
End Sub

Private Sub Command4_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command5_Click()
If List1.ListCount = 50 Then
MsgBox "You have 50 lines of text already! Please delete some!", vbCritical, "50 Lines of welcome text reached!"
Exit Sub
End If

Tmp = InputBox("Enter string to display when Profile is loaded. " & 50 - List1.ListCount & "/50 Remaining", "Add Line")
If Tmp > "" Then
List1.AddItem Tmp
End If

End Sub

Private Sub Command6_Click()
List1.Clear

End Sub


Private Sub Form_Load()
' Uses your current profile as a template,
' instead of you having to do your profile again

a1(0).BackColor = Heading
a1(1).BackColor = Msg
a1(2).BackColor = dis
a1(3).BackColor = con
a1(4).BackColor = svr
a1(5).BackColor = act
a1(8).BackColor = ThatsGood
a1(9).BackColor = ThatsBad
a1(10).BackColor = frmMain.BackColor
a1(11).BackColor = frmMain.txtSend.BackColor
a1(12).BackColor = frmMain.iChat.BackColor
a1(6).BackColor = frmMain.txtSend.ForeColor
Text1.Text = frmPic
Text2.Text = CaptionPrefix
'a1(7).BackColor = frmMain.cmdAction.BackColor
Text3.Text = ReadText("Theme", "Author", 1)
Text4.Text = ReadText("Theme", "Description", 1)
For i = 0 To 50
If ReadText("Theme", "WelcomeMSG" & i, 1) > "" Then List1.AddItem ReadText("Theme", "WelcomeMSG" & i, 1)
Next i
Text5.Text = EndWindow
End Sub

Private Sub Text1_DblClick()
' Loads a background.
' Not quite perfect, but it works ok
dlgColour.Filter = "Bitmap Files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg"
dlgColour.DialogTitle = "Select Background Picture..."
dlgColour.ShowOpen
Text1.Text = dlgColour.FileName

End Sub

Private Sub txt(Text As String, Optional Colour As ColorConstants, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As AlignmentConstants)
' Custom Text commands
' Like the Text Sub in module1, but much smaller

Text = Replace(Text, "+username+", UserName)
Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
Text = Replace(Text, "+room+", RoomName)

With frmNewProfile.iChat
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelBold = Bold
    .SelItalic = Italic
    .SelUnderline = Underline
    .SelFontSize = Size
    .SelAlignment = Alignment
    .SelColor = Colour
    .SelText = Text
    .SelLength = Len(.Text)
End With

End Sub

Private Sub Preview()
' Does what is says... duh! :)
SSTab1 = 2

iChat.Text = ""
iChat.BackColor = a1(12).BackColor

txt "Welcome to NChat ", Heading, True, False, False, 18, "Center"
txt "Alpha!!" & vbCrLf, ThatsGood, True, False, False, 18, "Center"
txt "Created by Grayda of Solid Inc." & vbCrLf, Heading, , , , 7, "Center"
txt ":nchat" & vbCrLf & "http://www.solidinc.tk" & vbCrLf, Heading, False, False, False, 7, "Center"
txt "Preparting to Start NChat Alpha, Please stand by..." & vbCrLf, Heading, True, False, False, 8, "Center"

txt Text3.Text & vbCrLf, a1(0).BackColor, True, , , , "Center"
For i = 0 To 49
If List1.List(i) > "" Then txt List1.List(i) & vbCrLf, a1(0).BackColor, True, False, True, , "Center"
Next i
txt "+username+ has entered the conversation!" & vbCrLf, a1(3).BackColor, True, False, False, 8, "Center"
txt "Welcome to room #" & frmMain.sckUDP.LocalPort & " +username+! I hope you enjoy your stay" & vbCrLf, a1(4).BackColor, True
txt "|| +username+ ||  G'day. Hows it going everyone?" & vbCrLf, a1(1).BackColor
txt "+username+ waves to everyone" & vbCrLf, a1(5).BackColor, False, True
txt "|| Some_Guy ||  Hey +username+" & vbCrLf, a1(1).BackColor
txt "|| +username+ ||  Hey Some_Guy." & vbCrLf, a1(1).BackColor
txt "An Error has occured." & vbCrLf, a1(9).BackColor, True
txt "Some_Guy has given you 50 NCredits!" & vbCrLf, a1(8).BackColor, True
txt "|| Some_Guy ||  I'm outta here, catch you later, k?" & vbCrLf, a1(1).BackColor
txt "Some_Guy has left the room" & vbCrLf, a1(2).BackColor, True, False, False, 8, "Center"
iChat.SelStart = 0
End Sub

