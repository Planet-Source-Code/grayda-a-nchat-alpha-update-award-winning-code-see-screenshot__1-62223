VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRooms 
   Caption         =   "List of Available Rooms"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   Icon            =   "frmRooms.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Join a custom room"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4215
   End
   Begin MSComctlLib.TreeView List1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6165
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Room Info"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Join Selected Room!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Menu mnuRC 
      Caption         =   "RightClick Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuJoin 
         Caption         =   "Join this room"
      End
      Begin VB.Menu hr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetInfo 
         Caption         =   "Get information about this room"
      End
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    List1_DblClick

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If List1.SelectedItem.Parent.Text = "NChat Rooms" Then
        Select Case List1.SelectedItem.Index
        Case "3"
            MsgBox "This is the default room for NChat. Most people choose to stay in here to chat", vbInformation, "The Lobby"
        Case "4"
            MsgBox "Music Chat room. You can come in here to discuss anything music related, such as Top 40 songs, or songs that you would like to get your hands on", vbInformation, "Music Chat"
        Case "5"
            MsgBox "Work? What's work? If you know, then you can come in here and talk about it", vbInformation, "The Work Room"
        Case "6"
            MsgBox "Help for NChat. If there is someone in there, then you can ask them for help about NChat (Just in case my help file wasn't enough >:( )", vbInformation, "Help for NChat!"
        Case "7"
            MsgBox "Programmers Chat. Come in here for help with programming in any language (Even Japanese!!... um... providing someone speaks Japanese)", vbInformation, "Programmers Chat"
        Case "8"
            MsgBox "Room for fighters. This is where you go to fight it out. The admin may re-direct you to this room to let you cool off", vbInformation, "Room For Fighters"
        End Select
        Exit Sub
    End If
    RoomBroadcast "infoø" & List1.SelectedItem.Text & "ø" & UserName

End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Tmp = InputBox("Please enter a room number to connect to", "Join Custom Room")
    If Tmp > 0 And Tmp <> 3155 And Tmp <> 1113 Then
        NewRoom Tmp, "Custom Room (Room " & Tmp & ")"
    Else
        MsgBox "Invalid room or blank! Cannot be 3155 or 1113", vbExclamation, "Bad Room Number"
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Move 150, 150, Me.Width - 500, Me.Height - Command1.Height - Command3.Height - 900
    Command1.Move 100, List1.Height + 350 + Command3.Height, List1.Width / 2
    Command2.Move Command1.Width + 200, Command1.Top, List1.Width / 2
    Command3.Move 100, List1.Height + 200, List1.Width
End Sub


Private Sub List1_DblClick()
    On Error GoTo Errs
    If CreatedRoom = True Then
AskSelect:
        If MsgBox("You are about to change rooms. This will cause your room not to be listed in the list of rooms. If you like, you can nominate a NEW room host, who will be able to set the welcome message, kick users etc. Do you want to nominate a NEW room host? If you click NO, then your room will disappear.", vbQuestion + vbYesNo, "Change rooms?") = vbYes Then
            NewHost = ShowBox("Select", "Select New Room Host")
            If NewHost = UserName Then
                MsgBox "You cannot select yourself! That is a security risk! Please choose a valid user or click Cancel", vbCritical, "Bad Room Host!"
                Exit Sub
            ElseIf NewHost = "" Then
                Exit Sub
            End If

            If NewHost > "" Then Broadcast "nrhø" & RoomName & "ø" & RoomID & "ø" & NewHost & "ø" & Description & "ø" & Password
        End If
    End If

    Dim Parse() As String

    Parse = SplitVB5(List1.SelectedItem.Key, "/ù/")
    If Parse(1) > "" Then
        If InputBox("This room is protected by a password. Please enter it to continue", "Password Protected Room!") <> Decode(Parse(1), "FireStormInc") Then
            MsgBox "Password Incorrect!"
            Exit Sub
        End If
    End If
skipo:
    If List1.SelectedItem.Text = "NChat Rooms" Or List1.SelectedItem.Text = "User-Made Rooms" Then Exit Sub

    If Right(Parse(0), Len(Parse(0)) - 1) > 0 And List1.SelectedItem.Text > "" Then
        NewRoom Right(Parse(0), Len(Parse(0)) - 1), List1.SelectedItem.Text
        Unload Me
    Else
        MsgBox "This room is missing connection info. Cannot connect!", vbCritical, "Missing Room Info!"
    End If
    Exit Sub
Errs:
    GoTo skipo
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And List1.SelectedItem.Text > "" Then PopupMenu mnuRC
End Sub

Private Sub mnuGetInfo_Click()
    Command2_Click
End Sub

Private Sub mnuJoin_Click()
    Command1_Click
End Sub
