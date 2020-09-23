VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAutobot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notch - The NChat Bot Machine"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6855
   Icon            =   "frmAutobot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NChat_Button3 
      Caption         =   "Notch Options"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog D1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton NChat_Button2 
      Caption         =   "Help"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton NChat_Button1 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton NChat_Button4 
      Caption         =   "Start Notch"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton NChat_Button5 
      Caption         =   "Hide Window"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton NChat_Button6 
      Caption         =   "Stop Notch"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton NChat_Button7 
      Caption         =   "Reload Notch"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Load a Notch File:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAutobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is for handling some AutoBot stuff.
' The editing of the script has been removed because
' of the complexity of the Notch Scripts

Private Sub NChat_Button1_Click()
' Opens the dialog box to let us select a script
    d1.Filter = "Notch INI File (*.ini)|*.ini|All Files (*.*)|*.*"
    d1.ShowOpen

    If d1.FileName > "" Then
        List1.AddItem "(" & Time & ") Notch File Loaded!"
        Text1.Text = d1.FileName
    End If

End Sub

Private Sub NChat_Button2_Click()
' Just some very simple help. There is an external
' help file for Notch in the <other stuff> folder
    Dim Help As String

    Help = Help & "What has happened to the Notch Script Editor?" & vbCrLf & vbCrLf & _
           "The Notch script editor has been removed, because of the New Notch script system. To code an editor will take advanced coding, which I don't have the time to make. Instead, all scripting is done through an INI file. For more help, consult the Notch Help file that is packaged with the source code" & vbCrLf & vbCrLf

    MsgBox Help


End Sub



Private Sub NChat_Button3_Click()
    frmAutoBotOptions.Show

End Sub

Private Sub NChat_Button4_Click()
' Starts Notch!
    Dim TMPA As String
    ' Missing file or no file? Then don't start Notch
    If Trim(Text1.Text) = "" Or FileObj.FileExists(Text1.Text) = False Then
        MsgBox "Cannot start Notch! File not found, or missing!", vbCritical, "Cannot Load!"
        Exit Sub
    End If

    ' Set the NotchRunning flag
    NotchRunning = True
    ' Tell our timer to accept times entered as seconds
    ' This timer has some problems accepting 90 seconds
    ' as an interval, and I'm not sure why!!
    frmAutoBotOptions.Timer1.Interval = Val(frmAutoBotOptions.Text2.Text) * 1000

    IniFile(2) = Text1.Text
    ' Our idle chat timer
    If frmAutoBotOptions.Check1.Value = 1 Then frmAutoBotOptions.Timer1.Enabled = True
    List1.AddItem "(" & Time & ") Notch system started"
    Dim F As Integer
    ' YOu can have upto 50 "Welcome" messages. These
    ' are run on Notch Startup, so you can add notch
    ' as a 'real' user, kick a user, display a message etc.
    F = 0
    For I = 1 To 50
        TMPA = ReadText("General", "Welcome" & I, 2)
        If Trim(TMPA) > "" Then
            F = F + 1
            Broadcast TMPA
            DoEvents
            DoEvents
            DoEvents
            DoEvents
        End If
    Next I

    ' Enumsect (Found in modAI) tells us how many
    ' questions Notch has in the script file
    List1.AddItem "(" & Time & ") " & EnumSect & " Phrases loaded"
    List1.AddItem "(" & Time & ") " & F & " Welcome lines Broadcasted"
    List1.AddItem "(" & Time & ") Notch is ready to interact!"
    ' Display a popup box. Not shown on < Windows 98
    Tray.Box "Notch has been activated, and is running in the room " & RoomName & ". To stop, open the Bot menu in the Admin menu", "Notch Started"

End Sub




Private Sub NChat_Button5_Click()
' Don't unload the form, because then our idle
' chatter engine won't work!
    Me.Visible = False

End Sub

Private Sub NChat_Button6_Click()
    Dim TMPA As String


    Tray.Box "Notch has been stopped at: " & Time & ". He learnt: " & NewWords & " New Words, and was stopped cleanly", "Notch Stopped"
    frmAutoBotOptions.Timer1.Enabled = False
    Dim F As Integer
    ' 50 more shutdown messages. Used to broadcast
    ' disconnect commands etc.
    F = 0
    For I = 1 To 50
        TMPA = ReadText("General", "Shutdown" & I, 2)
        If Trim(TMPA) > "" Then
            F = F + 1
            Broadcast TMPA
            DoEvents
            DoEvents
            DoEvents
            DoEvents
        End If
    Next I
    IniFile(2) = ""
    List1.AddItem "(" & Time & ") " & F & " Shutdown lines Broadcasted"
    List1.AddItem "(" & Time & ") Notch has been stopped."
    NotchRunning = False
End Sub

Private Sub NChat_Button7_Click()
' Simple reload of Notch. Doesn't re-broadcast
' the welcome messages
    IniFile(2) = Text1.Text
    List1.AddItem "(" & Time & ") Notch File has been reloaded.."
End Sub




