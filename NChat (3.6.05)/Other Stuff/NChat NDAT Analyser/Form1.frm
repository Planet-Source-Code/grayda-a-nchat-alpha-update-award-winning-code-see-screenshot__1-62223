VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NChat Data Decrypter v2.0"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgLoad 
      Left            =   5430
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "NChat Settings File (*.ndat)|*.ndat"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyse"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program lets you view, at a glance, your decrypted settings file.
' This lets you work out what settings you have enabled / disabled. It also
' lets you work out errors in your settings file, and determine why some features
' have not been loaded

' Gets the proper windows username, instead of using environ("username")
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUserName() As String
' Simple sub to get our windows username
   Dim UserName2 As String * 255
   Call GetUserNameA(UserName2, 255)
   GetUserName = Left$(UserName2, InStr(UserName2, Chr$(0)) - 1)
End Function

Private Sub LoadSettings()
On Error Resume Next
' Our settings to be applied
Dim Settings() As String
' Our encoded, then decoded settings file
Dim TheFile As String

' Open the file for reading (input)
Open dlgLoad.FileName For Input As #1
' Load it into our string, which can be of almost any length
TheFile = Input$(LOF(1), 1)
' Decode the file so it's plaintext
TheFile = Decode(TheFile, GetUserName & "€€NCHAT_SOLID¶¶" & GetUserName)
Close #1

' Split the file up according to our delimiter
Settings = Split(TheFile, "/©/")

' Check our settings file to make sure it's up-to-date
' FileCheck is a constant that tells us our settings file version
List1.AddItem "File Check-Code: " & Settings(0)

List1.AddItem "Your Username: " & Settings(1)

' This is how many icons you can pick from the list on frmOptions. YOu can buy more
List1.AddItem "Number of icons available: " & Settings(2)
' No icons to pick? Set them up with the default icons

If Settings(3) = "TharSheBlows" Then
List1.AddItem "Administrator?: False"
Else
List1.AddItem "Administrator?: True"
End If

List1.AddItem "Number of NCredits" & Settings(4)

' Username check stops cheating. NChat takes this setting (For example 'Billy') and
' compares it to your current windows username. If they are different, then you may
' have tried to dishonestly earn NCredits.
List1.AddItem "Anti-Cheat Username Check: " & Settings(5)


List1.AddItem "User Icon Name: " & Settings(6)

' Is swearing filtered out?
List1.AddItem "Swearing?: " & Settings(7)

' Your last profile loaded
List1.AddItem "Last profile loaded: " & Settings(8)

' True admins cannot be kicked from the NChat Chat room, or have their rights removed
List1.AddItem "Are you a true admin?: " & Settings(9)

List1.AddItem "Show tip of the day?: " & Settings(10)



List1.AddItem "Is your username bold?: " & Settings(11)
List1.AddItem "Is your username underlines?: " & Settings(12)
List1.AddItem "Your Username Text colour: " & Settings(13)

List1.AddItem "Message: " & Settings(14) & " " & GetUserName & " " & Settings(15)
List1.AddItem "Are smileys on?: " & Settings(16)
List1.AddItem "Are you banned from NChat?: " & Settings(17)

End Sub

Private Sub Command1_Click()


dlgLoad.ShowOpen
If dlgLoad.FileName = "" Then
Exit Sub
End If

LoadSettings
End Sub

Private Sub Command2_Click()
If dlgLoad.FileName = "" Then
MsgBox "Please load a file first!!", vbCritical, "No file loaded"
Exit Sub
End If
List1.Clear
LoadSettings
End Sub
