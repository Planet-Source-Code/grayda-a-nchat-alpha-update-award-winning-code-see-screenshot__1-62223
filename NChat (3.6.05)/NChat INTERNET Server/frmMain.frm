VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "// NChat Alpha Server v1.5 - Server OFFLINE \\"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1429
      ButtonWidth     =   1746
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            Key             =   "Stats"
            Object.ToolTipText     =   "// NChat Alpha Server Status \\"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            Key             =   "Users"
            Description     =   "Users"
            Object.ToolTipText     =   "// View All NChat Alpha Users \\"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Key             =   "Sep1"
            Description     =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop Server"
            Key             =   "StopServer"
            Description     =   "Stop"
            Object.ToolTipText     =   "// Stops the NChat Alpha Server \\"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start Server"
            Key             =   "StartServer"
            Description     =   "Start"
            Object.ToolTipText     =   "// Starts an NChat Alpha Server and lets people join \\"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblUptime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtLogs"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Timer1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "sckServer(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "List1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstUsers"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   -74880
         TabIndex        =   20
         Top             =   3840
         Width           =   8415
      End
      Begin MSComctlLib.ListView lstUsers 
         Height          =   5775
         Left            =   -69000
         TabIndex        =   19
         Top             =   1140
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   10186
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSWinsockLib.Winsock sckServer 
         Index           =   0
         Left            =   -71640
         Top             =   3480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   5513
         LocalPort       =   5513
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Server"
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   4500
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Server Details"
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   7335
         Begin VB.CommandButton Command5 
            Caption         =   "?"
            Height          =   255
            Left            =   6720
            TabIndex        =   17
            Top             =   1800
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "?"
            Height          =   255
            Left            =   6720
            TabIndex        =   16
            Top             =   1440
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "?"
            Height          =   735
            Left            =   6750
            TabIndex        =   15
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "?"
            Height          =   255
            Left            =   6750
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtServerName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   5055
         End
         Begin VB.TextBox txtServerDescription 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   1560
            TabIndex        =   7
            Top             =   600
            Width           =   5055
         End
         Begin VB.TextBox txtServerPassword 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1440
            Width           =   5055
         End
         Begin VB.TextBox txtServerMaxUsers 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   5
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Server Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Server Description:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Server Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Max Users Allowed:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -66840
         Top             =   3960
      End
      Begin RichTextLib.RichTextBox txtLogs 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   2
         Top             =   1080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0070
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -66840
         Top             =   6480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":00DE
               Key             =   "Status"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":09B8
               Key             =   "ConnectedUsers"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1692
               Key             =   "PowerOff"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":236C
               Key             =   "StartServer"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Online Users"
         Height          =   255
         Left            =   -69000
         TabIndex        =   18
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lblUptime 
         Caption         =   "Server Uptime: 00:00:00"
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   3480
         Width           =   2895
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

' This is the SERVER portion of NChat Alpha.
' You DO NOT need to run this if you are planning to
' use NChat over a LOCAL AREA NETWORK. ONLY if you
' plan to host an NChat room OVER THE INTERNET

' This is the simplist kind of server possible.
' Data comes in, the server sends it to all clients, who
' will decide if it's useful to them. Little parsing of data is done.
' The rest goes to clients. Understand? No? Damn...

' For setting our textbox to ONLY accept numbers
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const GWL_STYLE = (-16)
Const ES_NUMBER = &H2000&
' Is our server started?
Dim ServerActive As Boolean
' For Server Uptime
Dim Hours As Integer, Minutes As Integer, Seconds As Integer

Private Sub Command1_Click()

If ServerActive = False Then
' Start our server, and start the timer
Hours = 0
Minutes = 0
Seconds = 0
ServerActive = True

' Set our winsock up to listen for incoming requests
sckServer(0).Close
sckServer(0).LocalPort = 5133
sckServer(0).Listen
' Set our form caption up so we know our server is running
Me.Caption = "// NChat Alpha Server v1.5 - Server ONLINE \\"
' Log the server details, and default to tab #one (Tab 0), so you can't
' change the server details until you stop the current server
Text "Server Started successfully at: " & Now & vbCrLf, frmMain.txtLogs, DGreen, True
Text vbTab & "Server Name: " & txtServerName.Text & vbCrLf, frmMain.txtLogs, DGreen, True
Text vbTab & "Server Password: " & String(Len(txtServerPassword.Text), txtServerPassword.PasswordChar) & vbCrLf, frmMain.txtLogs, DGreen, True
Text vbTab & "Max Users: " & txtServerMaxUsers.Text & vbCrLf, frmMain.txtLogs, DGreen, True
Text vbTab & "Server IP: " & sckServer(0).LocalIP & vbCrLf, frmMain.txtLogs, DGreen, True
Text vbTab & "On Port: " & sckServer(0).LocalPort & vbCrLf, frmMain.txtLogs, DGreen, True
SSTab1.Tab = 0
Toolbar1.Buttons(5).Caption = "Edit Server Info"
Command1.Enabled = False
ElseIf ServerActive = True Then Exit Sub
End If
End Sub

Private Sub Command2_Click()
MsgBox "This is the NAME of your new Chat Room. If the name is: My Room, then when people connect to your room, then it will appear as: My Room", vbInformation, "Room Name"

End Sub

Private Sub Command3_Click()
MsgBox "This is the description of your room. What it does, what the topic is, who is allowed in and so forth. Keep it short and simple", vbInformation, "Server Description"
End Sub

Private Sub Command4_Click()
MsgBox "To restrict access to this chat room, you can provide a password. Give that password out to whoever you want, then only they can join. Good for private or top secret discussions", vbInformation, "Server Password"

End Sub

Private Sub Command5_Click()
MsgBox "If this is set to 0 or less, then you can have as many chatters as you want in your room. If it's greater than zero, then the number of people allowed to join the room is limited to that number", vbInformation, "Server Limit"


End Sub

Private Sub Form_Load()
' No server active. May code some command lines that will let you
' start the server from the command line, some day :)
ServerActive = False
' Enable the "Start Server" button
' but disable the "Stop Server" for obvious reasons
Toolbar1.Buttons(4).Enabled = False
Toolbar1.Buttons(5).Enabled = True

' Set our text box (For max number of connected users) to ONLY accept
' numbers and not letters. Thanks to AllApi.net for help with this!
     curstyle = GetWindowLong(txtServerMaxUsers.hwnd, GWL_STYLE)
    'Set the new style to ONLY accept Numbers
    SetWindowLong txtServerMaxUsers.hwnd, GWL_STYLE, curstyle Or ES_NUMBER
    'refresh the box
    txtServerMaxUsers.Refresh

' Log
Text "// NChat Alpha Server Loaded at: " & Time & " and is ready to start!" & vbCrLf, frmMain.txtLogs, Orange, True
SSTab1.Tab = 0
End Sub

Private Sub sckServer_Close(Index As Integer)
On Error Resume Next
' When a winsock closes, it's usually because the client has disconnected
' or the computer has shut down. In this case, remove the username
' from the list of users.
Text lstUsers.ListItems(Index).Text & " has left // NChat Alpha at: " & Time, frmMain.txtLogs, DRed, True
' Scroll through all winsocks until we find a connected one
For i = sckServer.LBound To sckServer.UBound
DoEvents
' Then Let the others know of their disconnection
If sckServer(i).State = 7 Then SendData "disø" & lstUsers.ListItems(Index).Text, i
Next i
' Finally remove their name from the server's list of users
lstUsers.ListItems.Remove (Index)

    
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
' TO DO: Insert Banned IP list for searching
Load sckServer(sckServer.UBound + 1)
' Is the new user pushing the server limit? Don't let them in!
For i = 1 To sckServer.UBound
    If sckServer(i).State <> 7 Then
        sckServer(i).Close
        sckServer(i).LocalPort = 0
        sckServer(i).Accept requestID

If sckServer.UBound = txtServerMaxUsers.Text And txtServerMaxUsers.Text > 0 Then
' Let them connect only to popup a message box then disconnect them
Text "ERROR: MAX USERS REACHED: " & txtServerMaxUsers.Text & ". TRY RAISING LIMIT" & vbCrLf, frmMain.txtLogs, vbRed, True
SendData "svrøCannot Connect to server. Server is full!", i
sckServer(i).Close
Exit Sub
End If
        
        DoEvents
        Text "[ " & Now & "] New User Connected: " & sckServer(i).RemoteHostIP & " " & sckServer.UBound & " / " & txtServerMaxUsers.Text & " Users connected!" & vbCrLf, frmMain.txtLogs, DGreen, True
        SendData "svrø" & txtServerName.Text & "+newline++newline+" & txtServerDescription.Text & "+newline++newline+You are user #" & i & " of " & txtServerMaxUsers.Text, i
        DoEvents
        DoEvents
        Exit Sub
     End If
   Next i



End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
' OK here is the gist of the DataArrival Sub:

' ALL Data is sent to the server. The server simply spits it back
' at everyone connected. The data is ONLY parsed by the server if it's
' a server only thing, like kick requests and so forth. It may also be
' parsed by the server if it's a msg act or Private message, and the
' host wants a transscript of what's going on in their room.

' Holds our split data
Dim Result() As String
Dim sData As String

sckServer(Index).GetData sData

Result = SplitVB5(sData, "ø")

Select Case Result(1)

Case "pwd" ' Checks if their password to connect is correct
' Invalid password? Disconnect them
If Result(2) <> txtServerPassword.Text Then sckServer(Index).Close

Case "msg" ' Message that is sent to ALL people connected

Text "(" & Now & ") [" & Result(2) & "] " & Result(3) & vbCrLf, frmMain.txtLogs, vbBlue, True

Case "act" ' A third-person winmx style action. Lets you act something out
Text "(" & Now & ") " & Result(2) & Result(3) & vbCrLf, frmMain.txtLogs, Orange, True, True
Case "svr"
Text "(" & Now & ") " & Result(2) & vbCrLf, frmMain.txtLogs, Purple, True

Case "dis"
Text sckServer(i).RemoteHostIP & " Disconnected at [ " & Now & " ] " & vbCrLf, frmMain.txtLogs, DGreen, True

sckServer(Index).Close

Case "usr"
For i = 1 To lstUsers.ListItems.Count
If lstUsers.ListItems(i).Text = Result(2) Then Exit Sub
Next i

lstUsers.ListItems.Add , , Result(2)

End Select

For i = sckServer.LBound To sckServer.UBound
If sckServer(i).State = 7 Then
' Text "[" & Result(1) & "] " & Result(2) & vbCrLf, frmMain.txtLogs, vbBlue, True
' Join the results together and send it off for the client to dissect
' As we have no use for it yet.
SendData JoinVB5(Result, "ø"), i
End If
Next i
' Add the data to the list box. But only if it's useful NChat data, not user info
If Result(1) <> "usr" Then List1.AddItem JoinVB5(Result, "ø")
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' Error with the winsock? Then log an error message
sckServer(Index).Close
Text "ERROR WITH WINSOCK #" & Index & vbCrLf & vbTab & Number & vbCrLf & UCase(Description) & vbCrLf, frmMain.txtLogs, vbRed, True

End Sub

Private Sub Timer1_Timer()

If ServerActive = False Then
Toolbar1.Buttons(4).Enabled = False
'Toolbar1.Buttons(5).Enabled = True
Exit Sub
Else
'Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(4).Enabled = True
End If
' Server timer that acts in hours, minutes and seconds, like a digital clock.
' This gives you an acurate representation of how long your server has been running
Seconds = Seconds + 1
If Seconds >= 60 Then
Minutes = Minutes + 1
Seconds = 0
End If

If Minutes >= 60 Then
Hours = Hours + 1
Minutes = 0
End If

lblUptime.Caption = "Server Uptime: " & Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00") & " - " & sckServer(sckServer.UBound).State
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' When you click on a button, all it does is change the current tab
' of SSTab1, which has been carefully hidden

' Find out what toolbar button we are clicking on
Select Case Button.Caption
' Server Status
Case "Status"
' Select Tab 0 (The first one)
SSTab1.Tab = 0
Case "Users"
SSTab1.Tab = 1
Case "Stop Server"


For i = sckServer.LBound To sckServer.UBound
If sckServer(i).State = 7 Then
' Tell everyone to connect to another room, coz this server is being stopped
SendData "N©H@-|-øsvrøThis NChat server is being stopped. Please connect to a different room!", i
Text "// NChat Alpha Server v1.0 - Server STOPPED at " & Now & " \\", frmMain.txtLogs, vbRed, True
End If
Next i

Toolbar1.Buttons(5).Caption = "Start Server"
Command1.Enabled = True
ServerActive = False

sckServer(0).Close
Me.Caption = "// NChat Alpha Server v1.0 - Server STOPPED at " & Now & " \\"

Case "Start Server"
SSTab1.Tab = 3

Case "Edit Server Info"
SSTab1.Tab = 3
End Select

End Sub
