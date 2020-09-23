VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NChat Options"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel!!"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   83
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmOptions.frx":1B7A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ImageCombo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "d1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":1B96
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Check3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Admin"
      TabPicture(2)   =   "frmOptions.frx":1BB2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin MSComDlg.CommonDialog d1 
         Left            =   5040
         Top             =   2460
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Select Custom Icon"
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "frmOptions.frx":1BCE
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73680
         TabIndex        =   21
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73680
         TabIndex        =   20
         Top             =   3120
         Width           =   3975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Check Password"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71160
         TabIndex        =   19
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Admin Logout"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72720
         TabIndex        =   18
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Preview"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2580
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmOptions.frx":1D45
         Left            =   120
         List            =   "frmOptions.frx":1D64
         TabIndex        =   6
         Top             =   2820
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOptions.frx":1D8B
         Left            =   120
         List            =   "frmOptions.frx":1DAA
         TabIndex        =   5
         Top             =   2460
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Allow swearing?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reset NChat"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   3660
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Allow Smileys?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1860
         Width           =   1455
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   1500
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "Please Select an Icon"
      End
      Begin VB.Label Label7 
         Caption         =   "Username:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Password:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "This box allows you to change how messages look. Select or type one in both of the lists, and click PREVIEW"
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   1980
         Width           =   4650
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2940
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "This box allows you to change your username for free!!"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "In this box, you can select your icon. This wil appear next to your username on the list to the right."
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "If you are easily offended by foul language, NChat can block some language. To do so, click this box below"
         Height          =   495
         Left            =   -74880
         TabIndex        =   13
         Top             =   540
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "If you have a slow computer, or would like to disable pictures (Smileys), then uncheck the box below"
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   1380
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "If you want to start again, with NO settings file, NCredits etc, then click this button"
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   3180
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Your old username, just for broadcasting
' so it will say oldusername is now known as username
Dim OldUN As String

Private Sub Command7_Click()
    d1.Filter = "JPEG Files (*.jpg)|*.jpg|Windows Bitmap (*.bmp)|*.bmp|All files (*.*)|*.*"
    d1.DialogTitle = "Select custom icon to use"
    d1.ShowOpen
    Set frmMain.List1.SmallIcons = Nothing
    Set ImageCombo1.ImageList = Nothing

    FileCopy d1.FileName, AppPath & "\User Icons\" & d1.FileTitle
    frmMain.ImageList1.ListImages.Add 1, d1.FileTitle, LoadPicture(AppPath & "\User Icons\" & d1.FileTitle)
    Set frmMain.List1.SmallIcons = frmMain.ImageList1
    Set ImageCombo1.ImageList = frmMain.ImageList1
    ImageCombo1.ComboItems.Clear


    For I = 1 To TotalIcons
        ImageCombo1.ComboItems.Add , "K" & I, frmMain.ImageList1.ListImages.Item(I).Key, frmMain.ImageList1.ListImages.Item(I).Key
    Next I

    Broadcast "chiø" & MyIcon & "ø" & IP
End Sub

Private Sub Form_Load()
' Selects the first tab, instead of the IDE
' defined tab
    SSTab1.Tab = 0
    ImageCombo1.ImageList = frmMain.ImageList1
    ' Loads your options such
    ' as swearing and your icon
    'On Error Resume Next

    ' If you are an admin, then you can select all the
    ' icons, if not, then you need to buy them :)
    If frmMain.mnuDev.Visible = False Then
        For I = 1 To TotalIcons
            ImageCombo1.ComboItems.Add , "K" & I, frmMain.ImageList1.ListImages.Item(I).Key, frmMain.ImageList1.ListImages.Item(I).Key
        Next I
    Else
        For I = 1 To frmMain.ImageList1.ListImages.Count - 2
            ' Prefix the key with a "K" because keys can't be numbers only
            ImageCombo1.ComboItems.Add , "K" & I, frmMain.ImageList1.ListImages.Item(I).Key, frmMain.ImageList1.ListImages.Item(I).Key
        Next I
    End If
    ' Selects your current icon
    ImageCombo1.ComboItems.Item(FindIcon(MyIcon)).Selected = True

    ' If you are an admin, then show the logout button
    If frmMain.mnuDev.Visible = True Then Command5.Enabled = True

    Text1.Text = UserName
    ' Old UN lets NChat send messages like: OldUsername
    ' is now known as NewUsername and stuff
    OldUN = UserName
    ImageCombo1.ImageList = frmMain.ImageList1
    Combo1.Text = StartMSG
    Combo2.Text = EndMSG

    If Swearing = True Then Check1.Value = Checked
    If Smiley = True Then Check3.Value = 1

End Sub

Private Sub Command1_Click()
' If the checks are enabled, then
' turn the booleans on

'On Error GoTo Errors


' Ensure our imagelist is initialized.
    ImageCombo1.ImageList = frmMain.ImageList1


    ' CHecks your name for invalid stuff
    If Right(Trim(Text1.Text), 3) = "[A]" And frmMain.mnuDev.Visible = False Or Right(Trim(Text1.Text), 4) = "[RA]" And frmMain.mnuDev.Visible = False Then
        MsgBox "Sorry, but only administrators can have [A] or [RA] on the end of their name...", vbExclamation, "Bad Username"
        SSTab1.Tab = 0
        Exit Sub
    End If

    For I = 1 To frmMain.List1.ListItems.Count
        If Trim(frmMain.List1.ListItems.Item(I).Text) = Trim(Text1.Text) And Trim(frmMain.List1.ListItems.Item(I).Text) <> UserName Then
            MsgBox "Sorry, that username has been taken. Please select another one...", vbExclamation, "Username in use"
            SSTab1.Tab = 0
            Exit Sub
        End If
    Next I


    If Len(Text1.Text) > 25 Then
        MsgBox "Your username is too long! Please shorten it (25 letters max)", vbCritical, "Username too long!"
        SSTab1.Tab = 0
        Exit Sub
    End If
    OldUsername = UserName
    UserName = Trim(Text1.Text)

    If UserName = OldUsername Then
        UserName = OldUsername
    Else


        Broadcast "chuø" & OldUN & "ø" & MyIcon

        DoEvents
        ' Server Text syntax: svr <delimiter> Text to send
        Broadcast "svrø" & OldUsername & " is now known as " & UserName
        Text "Changed your Username from " & OldUN & " to " & UserName & "!" & vbCrLf, Heading, True, , , , "Center"
    End If

    If Check1.Value = 1 Then
        Swearing = True
    Else
        Swearing = False
    End If


    StartMSG = Combo1.Text
    EndMSG = Combo2.Text

    ' Sets MyIcon to imagecombo's icon
    If ImageCombo1.SelectedItem.Index <> FindIcon(MyIcon) Then
        Broadcast "chiø" & ImageCombo1.SelectedItem.Text
        MyIcon = ImageCombo1.SelectedItem.Text
        frmMain.List1.ListItems(1).SmallIcon = frmMain.ImageList1.ListImages(FindIcon(MyIcon)).Key
    End If

    If Check3.Value = Checked Then
        Smiley = True
    Else
        Smiley = False
    End If



    Unload Me
    Exit Sub
Errors:
    MyIcon = MyIcon
    Unload Me
End Sub


Private Sub Command2_Click()
' Cancel Button
    Unload Me
End Sub

Private Sub Command3_Click()
' Resets NChat so you can start again
    If MsgBox("YOU ARE ABOUT TO DELETE YOUR NCHAT SETTINGS. ARE YOU SURE YOU WANT TO DO THIS?", vbCritical + vbYesNo, "WARNING:") = vbYes Then
        Kill AppPath & GetUserName & ".ndat"
        MsgBox "NChat settings files deleted. NChat will now reset so you can start again", vbInformation, "Delete successful"
        End
    Else
        MsgBox "Delete aborted. No files have been deleted", vbCritical, "NO files deleted"
    End If
End Sub

Private Sub Command4_Click()

    If Text2.Text = UserName & "101" Then

        Enc = ""

        Enc = SHAHash("Thisismycodetohash" & GetUserName & "Ó¢œÛq+¦mJ--www.solidinc.tk" & GetUserName)

        If Text3.Text = Enc Then
            frmMain.mnuDev.Visible = True
            TrueAdmin = True
            Text2.Text = ""
            Text3.Text = ""
            MsgBox "Password Correct!!", vbInformation, "Administrator Login"
            Broadcast "svrø" & UserName & " is now an NChat Admin!"
            OLD = UserName
            UserName = Replace(UserName, " [A]", "")
            UserName = UserName & " [A]"
            Broadcast "chuø" & OLD & "ø" & MyIcon
            Unload Me
        Else
            MsgBox "Incorrect Password!!", vbCritical, "Wrong"
            Text2.Text = ""
            Text3.Text = ""
            Exit Sub
        End If
    Else
        MsgBox "Incorrect Username!!", vbCritical, "Wrong"
        Text2.Text = ""
        Text3.Text = ""
        Exit Sub
    End If
End Sub

Private Sub Command5_Click()
    frmMain.mnuDev.Visible = False
    TrueAdmin = False
    OLD = UserName
    UserName = Replace(UserName, " [A]", "")
    Broadcast "chuø" & OLD & "ø" & MyIcon
    Command5.Enabled = False
    Broadcast "svrø+username+ is no longer an admin!"
End Sub

Private Sub Command6_Click()
    Label10.Caption = Combo1.Text & " " & UserName & " " & Combo2.Text & " Hello!"
End Sub


Private Sub Text2_Change()
    If Text2.Text > "" And Text3.Text > "" Then
        Command4.Enabled = True
    Else
        Command4.Enabled = False
    End If

End Sub

Private Sub Text3_Change()
    If Text3.Text > "" And Text2.Text > "" Then
        Command4.Enabled = True
    Else
        Command4.Enabled = False
    End If

End Sub

