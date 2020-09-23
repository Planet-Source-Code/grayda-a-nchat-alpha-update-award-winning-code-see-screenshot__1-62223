VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NChat Administrator Password Generator"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy to clipboard"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   25
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program will generate your NChat Administrator Password.
' It has been designed ONLY for administrators, and is not meant
' to get into the hands of others.

' This password changes every hour. Therefore it ensures that people cannot
' use the same password twice (unless they know their password changes, and
' enter the password at the same time of day)

' This is the main part of our password encryption. This is what ensures our
' password is different for each person
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

' How long you have to enter your password
Dim Remaining As Integer
' When your password will expire
Dim Expire As String

Private Sub Command1_Click()

' How much longer you have to enter your password into NChat
Remaining = -(Minute(Time) - 60)

If Text1.Text = "" Then
MsgBox "You didn't enter a username! Please enter one now...", vbCritical, "No username"
Exit Sub
End If

If Len(Text1.Text) > 25 Then
MsgBox "Your username is too long! The password generated will not be compatible with NChat!! Please check the name and try again", vbCritical, "Username too long!"
End If

' The heart of our password obscurer. It basically takes each letter,
' adds the hour, and minute and username to it, makes it into an ascii code, makes
' it into hex, then spits out the response.
Text2.Text = SHAHash("Thisismycodetohash" & GetUserName & "Ó¢œÛq+¦mJ--www.solidinc.tk" & GetUserName)

MsgBox "Password Generated!! Switch to NChat, and click 'Main Menu', 'NChat Options', then 'Admin'. You then need to enter the following Information: " & vbCrLf & vbCrLf & "Username: " & Text1.Text & "101" & vbCrLf & "Password: " & Enc & vbCrLf & vbCrLf & "PS, you only have " & Remaining & " minutes to enter your password", vbInformation, "Done!"

End Sub

Private Sub Command2_Click()
' Copy to clipboard
If Text2.Text > "" Then
Clipboard.SetText Text2.Text
Else
MsgBox "Nothing to copy! Please generate a password First!!", vbCritical, "Nothing to copy!"
End If

End Sub

Private Sub Command3_Click()
End

End Sub

Private Sub Form_Load()
' Work out when our password will expire
If Hour(Time) > 12 Then
Expire = Hour(Time) - 12 + 1 & ":00 PM"
Else
Expire = Hour(Time) + 1 & ":00 AM"
End If


MsgBox "This program will generate your NChat Administrator's Password. To use it, type in your CURRENT NChat UserName (EXACTLY as it appears in NChat), and click 'Generate'. Your password will be valid until " & Expire & " and can be ONLY user on THIS computer. To generate passwords for other people, run this program on their computer", vbInformation, "PLEASE READ CAREFULLY"

End Sub

Private Function GetUserName() As String
' Simple function to retrieve a username
   Dim UserName2 As String * 255
   Call GetUserNameA(UserName2, 255)
   GetUserName = Left$(UserName2, InStr(UserName2, Chr$(0)) - 1)
End Function
