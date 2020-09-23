VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NChat - Private Chat - "
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1B7A
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   4560
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Label1 
      Caption         =   "Send!!"
      Default         =   -1  'True
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   3495
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5625
      Left            =   8355
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   9922
      ButtonWidth     =   1005
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            Object.ToolTipText     =   "Change the size of the line used for drawing"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Colour"
            Object.ToolTipText     =   "Changes the colour of the line being drawn"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Object.ToolTipText     =   "Clears the current drawing on both sides"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flood"
            Object.ToolTipText     =   "Flood-Fill a portion of the picture using the current line colour"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save the current images to your computer to keep"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2330
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26CA
            Key             =   "Flood"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A64
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   6360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrawType As String



Private Sub Form_Load()
' Can be draw or fill, depending on what you want to do
    DrawType = "Draw"

    ' Brings this window more in line with your profile
    'Text1.BackColor = frmMain.txtSend.BackColor
    'Text2.BackColor = frmMain.txtSend.BackColor
    'Me.BackColor = frmMain.BackColor
    ' Messages that you send are blue (by default)
    ' And incomming messages are orange (by default)
    'Text2.ForeColor = Msg


End Sub

Private Sub Form_Unload(Cancel As Integer)
' Frees up the window for the next chatter
    Me.Tag = ""
End Sub
Private Sub Label1_Click()
' Send the PM1 command
    If Trim(Text2.Text) > "" Then
        Broadcast "pm1ø" & Me.Tag & "ø" & Text2.Text & "ø" & UserName
        Txt1 UserName & " ::  " & Text2.Text & vbCrLf, Msg

        Text2.Text = ""
        ' Give em 2 NCredits... :)
        NCredits = NCredits + 2
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Flood fill the picture
    If Button = 1 And DrawType = "Flood" Then
        lastX = x
        lastY = y
        OldFill = Picture1.FillColor
        Picture1.FillColor = Col
        ' quite a simple API to use and call. This
        ' handles flood filling of a HDC

        ' The last draw code couldn't handle flood fills,
        ' because it used a VERY slow and resource consuming
        ' process where a new line (control) was created whenever
        ' you clicked. This was removed and updated by me!
        ExtFloodFill Picture1.hdc, x, y, Picture1.Point(x, y), 1
        Picture1.FillColor = OldFill

        Broadcast "fillø" & Me.Tag & "ø" & x & "ø" & y & "ø" & Col
    End If
End Sub





Private Sub Text1_Change()

' Not trying to copy anything? then scroll to the end
    If Text1.SelLength = 0 Then Text1.SelStart = Len(Text1.Text)
    'SetAutoURL4RTB Text1
End Sub

Public Sub Txt1(Text As String, Colour As String)
' Shoves the text into the textbox
    Text = Replace(Text, "+username+", UserName)
    Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
    Text = Replace(Text, "+room+", RoomName)
    Text = Replace(Text, "+ncredits+", NCredits)

    With Text1
        .SelStart = Len(.Text)
        .SelLength = Len(.Text)
        .SelColor = Colour
        .SelText = Text
        .SelLength = 0
    End With
    ' Replace the :) etc with the picture

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' This code was written by me.
' Build 7 and 8 used an old drawing style
' as is shown below. I did some planning,
' and have come up with a reasonably fast
' drawing code, which supports flood fills
' and other drawing APIs

' Here is the old code. ugh!
' Load a new line (This thing uses a hell of a lot of lines
' for complex drawings, so make sure to clear regularly
'Load Line1(Line1.Count)
' Set the new line at the cursor pos (lastX)
'Line1(Line1.UBound).X1 = lastX
' To the new X
'Line1(Line1.UBound).X2 = X
' And the same as above, but for Y
'Line1(Line1.UBound).Y1 = lastY
'Line1(Line1.UBound).Y2 = Y
'Line1(Line1.UBound).Visible = True

' Only draw if a button is pressed
    If Button = 1 And DrawType = "Draw" Then

        ' Tell the remote side you have moved the cursor
        Broadcast "moveø" & Me.Tag & "ø" & x & "ø" & y & "ø" & Col & "ø" & Picture1.DrawWidth
        Picture1.Line (lastX, lastY)-(x, y), Col
    End If

    ' Without this, when a picture is being drawn, and you
    ' hover over it, then it will distort something shocking!
    If Picture1.Enabled = True Then
        lastX = x
        lastY = y
    End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' Simple case stuff
    Select Case Button.Caption
    Case "Size"
        ' Change the size of the drawing line. New size is sent when
        ' you next draw a line
        Tmp = InputBox("New Line Width (1-50)", "New Line Width", Picture1.DrawWidth)

        If Tmp > 0 And Tmp < 51 Then
            Picture1.DrawWidth = Tmp
        Else
            MsgBox "Line width MUST be between 1 and 50!", vbCritical, "Too large or too small"
        End If

        ' Change the colour. Same as above
    Case "Colour"
        d1.ShowColor
        'Broadcast "colourø" & Me.Tag & "ø" & UserName & "ø" & d1.Color
        Col = d1.color

        ' Clears the line
    Case "Clear"
        If MsgBox("Are you sure you want to clear the whiteboard?", vbQuestion + vbYesNo, "Clear") = vbYes Then

            Broadcast "clearø" & Me.Tag & "ø" & UserName
            Picture1.Cls

        End If

    Case "Flood"
        If DrawType = "Draw" Then
            DrawType = "Flood"
            Toolbar1.Buttons.Item(Button.Index).Value = tbrPressed
        Else
            DrawType = "Draw"
            Me.Toolbar1.Buttons.Item(Button.Index).Value = tbrUnpressed
        End If

    Case "Save"
        d1.ShowSave
        SavePicture Picture1.image, d1.FileName

    End Select

End Sub
