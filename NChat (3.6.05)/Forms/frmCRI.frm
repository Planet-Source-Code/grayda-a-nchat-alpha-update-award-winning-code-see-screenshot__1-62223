VERSION 5.00
Begin VB.Form frmCRI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a new room"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmCRI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "?"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel!!"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Announce my room"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "Room Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Room Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Welcome message:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Your room's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCRI.frx":1B7A
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmCRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form lets you create your own NChat room.

Private Sub Command1_Click()
' Help button
    MsgBox "This is the message that everyone will see when they enter the room. It appears in the form of a purple (server) message. To use the person's name (For example: Welcome to " & UserName & "'s room Billy), then use +username+ (With the plusses)", vbInformation, "Welcome Message"
End Sub

Private Sub Command2_Click()
' Help button
    MsgBox "This is your room description (or topic). This will let people know what your room is about. It can be anything, like 'My room for General Chatting'", vbInformation, "Room Description"
End Sub

Private Sub Command3_Click()
' Help button
    MsgBox "If this is checked, then everyone (Even in other rooms) will be notified when your room is created", vbInformation, "Announce my room"
End Sub

Private Sub Command4_Click()
' Set up the random number generator
    Randomize

    ' Check for missing info
    ' None? Then set the new room name as text1
    If Text1.Text > "" Then
        RoomName = Text1.Text

    Else
        MsgBox "You didn't set a room name!", vbExclamation, "Incorrect Information!"
        Exit Sub
    End If

    ' Your room description (Topic)
    If Text3.Text > "" Then
        Tmp = Text3.Text
    Else
        MsgBox "You didn't set a description for your room!", vbExclamation, "Incorrect Information!"
        Exit Sub
    End If

    ' Pick your new room number. Not sure if this will conflict with other
    ' UDP-Enabled applications, but who cares? :P
    NewRoom Int(Rnd * 10000 + 255), Text1.Text
    ' the com command is only sent through sckRooms, and lets you send messages
    ' to EVERYONE connected to NChat
    If Check1.Value = 1 Then RoomBroadcast "comøA new room has been created by " & UserName & " called " & RoomName

    ' No NChat data (ie. msg, con, dis etc.)? Then set the
    ' welcome message as a server (svr) message

    If Left(Text2.Text, 3) <> svr Then
        WelcomeMsg = "svrø" & Text2.Text
    Else
        WelcomeMsg = Text2.Text
    End If
    If Text4.Text > "" Then Password = Text4.Text
    Description = Tmp
    Text "Room Created! Other people can join your room by clicking -Chat Rooms-, -List all Rooms- and then finding: " & RoomName & vbCrLf, "ThatsGood", True

    DoEvents
    CreatedRoom = True
    Unload Me

End Sub

Private Sub Command5_Click()
    Unload Me

End Sub

Private Sub Command6_Click()
    MsgBox "THIS FEATURE IS OPTIONAL. If you enter something into this box (Even spaces), then your room will be password protected. If it's password protected, then only people who know the password will be allowed to enter", vbQuestion, "Password"

End Sub

Private Sub Command8_Click()
' Help button
    MsgBox "This is the name for your room. Keep it simple, like " & UserName & "'s room", vbInformation, "Room Name"

End Sub

Private Sub Form_Load()
' Default room name
    Text1.Text = UserName & "'s Room"
End Sub
