VERSION 5.00
Begin VB.Form frmServerConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to NChat Server"
   ClientHeight    =   2070
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "frmServerConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   840
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   1680
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "5133"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect!"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Server Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Connection progress: "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Port (Optional):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Server IP or Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmServerConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Has our server password been sent yet?
Dim SendPWD As Boolean

Private Sub Command1_Click()
    IsInternet = True
    Address = Text1.Text
    frmMain.sckUDP.Close
    'frmMain.sckRooms.Protocol = sckTCPProtocol
    frmMain.sckUDP.Protocol = sckTCPProtocol
    frmMain.sckUDP.LocalPort = Text2.Text
    frmMain.sckUDP.RemotePort = Text2.Text
    frmMain.sckUDP.Connect Address
    'frmMain.sckRooms.Close
    'frmMain.sckRooms.Connect Text1.Text, 127

    frmMain.mnuChatRooms.Visible = False

End Sub

Private Sub Timer1_Timer()

    Label3.Caption = "Connection progress: " & frmMain.sckUDP.State
    'If frmMain.sckUDP.State = 7 And SendPWD = False Then
    'Broadcast "pwdø" & Text3.Text
    'SendPWD = True
    'End If

End Sub
