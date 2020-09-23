VERSION 5.00
Begin VB.Form frmCheating 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WARNING: CHEATING DETECTED"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmCheating.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   $"frmCheating.frx":1B7A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   5970
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "3) Your settings file has become corrupted and useless"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   5700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2) You have changed your Windows Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "1) You have copied someone else's NChat settings file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   5670
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NChat's security system has detected that you are using an NChat settings file that is not your own. This can happen by 3 methods:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This NChat Settings file does not belong to you..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmCheating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OK this form tells you when you have been caught
' trying to copy an information (ndat) file that isn't
' yours. Not the most secure method, but it works
Private Sub Command1_Click()
    Unload Me
End Sub

