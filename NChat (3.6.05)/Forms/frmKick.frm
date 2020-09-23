VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kick a user from NChat"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4635
   Icon            =   "frmKick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KICK USER"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7011
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmKick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A lot of people were complaining that they couldn't kick people because of
' special characters in their username, and couldn't copy and paste, so this
' form makes it easier

Private Sub Command1_Click()
    SelUser = ListView1.SelectedItem.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    SelUser = ""
    Unload Me


End Sub

Private Sub Form_Load()
' Load the list of users from the main form
    ListView1.ListItems.Clear
    ListView1.SmallIcons = frmMain.ImageList1.Object
    For I = 1 To frmMain.List1.ListItems.Count
        ListView1.ListItems.Add , frmMain.List1.ListItems.Item(I).Text, frmMain.List1.ListItems.Item(I).Text, , frmMain.List1.ListItems.Item(I).SmallIcon
    Next I
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem.Index > 0 Then Command1_Click

End Sub
