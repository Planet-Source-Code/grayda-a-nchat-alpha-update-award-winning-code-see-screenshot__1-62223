VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NChat Alpha Smiley Packer. Make the smileys YOU want!"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Step one: Select the image to use"
      Filter          =   $"Form1.frx":0000
      Flags           =   4
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Compile NChat Smiley Pack!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit current smiley"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove current Smiley"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add new smiley"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Current Smileys"
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   495
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   495
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   495
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   4575
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Showcase your NChat smileys on www.solidinc.tk in the NChat forums, and maybe your profile will become an official one!"
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":009D
      Height          =   1455
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This stuff is for displaying that really cool "Browse for Folder" window
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
' End "Browse for Folder" window declarations, const and types

Private Sub Command1_Click()
CD1.ShowOpen

If CD1.FileName = "" Then
MsgBox "No file selected! To add a smiley, you MUST add a valid picture", vbCritical, "No picture / invalid picture selected"
Exit Sub
End If

SmileyCode = InputBox("Please enter the smiley code. This code determines when the smiley appears (For example when you type :afro or :) or :( )", "Enter Smiley Code")

If Trim(SmileyCode) = "" Then
MsgBox "No Smiley Code entered! To add a smiley, you MUST add a valid Smiley Code", vbCritical, "No Smiley Code entered!"
Exit Sub
End If

Caption = InputBox("Please enter the smiley CAPTION. This appears when you hover your mouse over the smiley. This CAN be blank if required", "Enter Smiley CAPTION")

List1.AddItem SmileyCode
List2.AddItem CD1.FileName
List3.AddItem Caption
List4.AddItem CD1.FileTitle
End Sub

Private Sub Command2_Click()
On Error Resume Next
List2.RemoveItem List1.ListIndex
List3.RemoveItem List1.ListIndex
List4.RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex

End Sub

Private Sub Command3_Click()
On Error Resume Next
CD1.ShowOpen

If CD1.FileName = "" Then
MsgBox "No file selected! To add a smiley, you MUST add a valid picture", vbCritical, "No picture / invalid picture selected"
Exit Sub
End If

SmileyCode = InputBox("Please enter the smiley code. This code determines when the smiley appears (For example when you type :afro or :) or :( )", "Enter Smiley Code", List1.List(List1.ListIndex))

If Trim(SmileyCode) = "" Then
MsgBox "No Smiley Code entered! To add a smiley, you MUST add a valid Smiley Code", vbCritical, "No Smiley Code entered!"
Exit Sub
End If

Caption = InputBox("Please enter the smiley CAPTION. This appears when you hover your mouse over the smiley. This CAN be blank if required", "Enter Smiley CAPTION", List3.List(List1.ListIndex))

List1.List(List1.ListIndex) = SmileyCode
List2.List(List1.ListIndex) = CD1.FileName
List3.List(List1.ListIndex) = Caption
List4.List(List1.ListIndex) = CD1.FileTitle

End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure you want to compile this smiley pack? All selected images will be COPIED to the pack's folder (Which you can pick in the next step)? Clicking NO will stop the process", vbQuestion + vbYesNo, "Continue?") = vbNo Then Exit Sub

temp = BrowseForFolder
If temp = "" Then Exit Sub
If Right(temp, 1) <> "\" Then temp = temp & "\"

Open temp & "smiley.SMI" For Output As #1

For i = 0 To List1.ListCount - 1
Print #1, "[" & List1.List(i) & "]"
Print #1, "Location=" & List4.List(i)
FileCopy List2.List(i), temp & List4.List(i)
Print #1, "Caption=" & List3.List(i)
Print #1, ""
Next i
Close #1

MsgBox "Your smiley pack has been compiled to: " & temp & ". All you need to do now is zip it up using winzip, then it's ready to distribute! Or you can just copy the smileys and the smileys file to the NChat smileys folder to use them immediately!", vbInformation, "Done!"

End Sub

' Lets us Browse for a folder using the Windows API call
Public Function BrowseForFolder() As String
  Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = Me.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(AppPath, "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

BrowseForFolder = sPath

End Function


