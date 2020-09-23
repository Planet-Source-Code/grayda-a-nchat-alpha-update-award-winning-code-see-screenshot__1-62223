VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NChat Generation 2 Tip of the day"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6195
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't show again"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous Tip"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Tip"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frmTip.frx":1B7A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tip of the day goes here. It is loaded from the .RES file"
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Current Tip displaying
Dim CurTip As Integer

Private Sub Command1_Click()
    On Error Resume Next
    CurTip = CurTip + 1
    If CurTip = 17 Then CurTip = 1

    Label3.Caption = LoadResString(CurTip)
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    CurTip = CurTip - 1
    If CurTip = 0 Then CurTip = 16

    Label3.Caption = LoadResString(CurTip)

End Sub

Private Sub Command3_Click()
    If Check1.Value = 1 Then
        DontShowTip = True
    Else
        DontShowTip = False
    End If

    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next
    OnTop Me.hwnd
    Randomize
    If DontShowTip = True Then Check1.Value = 1
    CurTip = Int(Rnd * 16) + 1
    Label3.Caption = LoadResString(CurTip)
End Sub

