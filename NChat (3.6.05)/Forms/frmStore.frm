VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to the NChat Store! - 100 Credits remaining"
   ClientHeight    =   5400
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6780
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStore.frx":1B7A
            Key             =   "Cash"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStore.frx":1F60
            Key             =   "Basket"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmStore.frx":242F
      Top             =   1320
      Width           =   2295
   End
   Begin MSComctlLib.TreeView tvItems 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9128
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Purchase this item"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Item Information"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "This item costs: 0 NCredits"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
' UPDATE: Woah! Total Item Re-Design here.
' More items added, some removed, some modified...
' If you have any more ideas for NChat items,
' then send them in! firestorm_visual@hotmail.com
Dim Splice() As String


Private Sub cmdBuy_Click()
    On Error Resume Next

    Splice = SplitVB5(tvItems.SelectedItem.Key, "|")

    If tvItems.SelectedItem.Index > 0 Then
        Cost = Mid(Splice(0), 2, Len(Splice(0)) - 3)
    Else
        MsgBox "Please select an Item!!", vbCritical, "No Item Selected!"
        Exit Sub
    End If

    If Cost > NCredits Then
        MsgBox "Dude, check out the price of '" & tvItems.SelectedItem.Text & "' (" & Cost & " NCredits)! You only have " & NCredits & " NCredits Remaining...", vbCritical, "Out of funds"
        Exit Sub
    End If


    Select Case Splice(0)
    Case "H1", "H2", "H3", "H4"
        MsgBox "Please select an Item!!", vbCritical, "No Item Selected!"

        Exit Sub

    Case "N50AB"
        Broadcast "svrø" & InputBox("Please enter a message to send. It will be sent as a server (Purple Message)", "Send Purple Message")
    Case "N150CD"
        Broadcast "fakø" & InputBox("Enter who the message is from", "Send Fake Message") & "ø" & InputBox("Enter the fake message", "Send Fake Message")
    Case "N100EF"
        Broadcast "msgøAnonymous Userø" & InputBox("Enter the message to send as an anonymous user", "Send anonymous user")
    Case "N200GH"
        Broadcast "heaø" & InputBox("Enter message to send as a large (Heading) message", "Send Heading")
    Case "N150IJ"
        Broadcast "actø" & InputBox("Enter username to send action as", "Send Fake Action") & "ø" & InputBox("Enter action to send", "Send Fake Action")
    Case "N1000KL"
        RoomBroadcast "comø" & InputBox("Enter message to send. It will be sent to EVERYONE in NChat!", "Send message to EVERYONE")
    Case "N300MN"

        Randomize
        Decider = Int(Rnd * 50)
        If Decider <= 25 Then
            Change = Int(Rnd * 500)
            NCredits = NCredits + Change
            Text "Added " & Change & " to total NCredits!!" & vbCrLf, "ThatsGood", False, True
        ElseIf Decider > 26 Then
            Change = Int(Rnd * 500)
            NCredits = NCredits - Change
            Text "Subtracted " & Change & " from total NCredits!!" & vbCrLf, "ThatsBad", False, True
        End If

    Case "N100OP"

        Randomize
        Tmp = InputBox("Enter a lucky number between 1 and 100. If you get it right, then you will win 10000 NCredits", "Lucky Number")
        Num = Int(Rnd * 100)
        If Num = Tmp Then
            NCredits = NCredits + 10000
            MsgBox "Congratulations! You win 10000 NCredits!", vbExclamation, "YOU WIN!!"
        Else
            MsgBox "Bad luck. Maybe next time (The lucky number was " & Num & ")", vbInformation, "Bad luck"
        End If

    Case "N1000QR"
        ShowBox "Steal NCredits", "Steal NCredits from user"
        If SelUser = "" Then Exit Sub
        Broadcast "sndø" & SelUser & "ø" & -Int(Rnd * 2000)

    Case "N500ST"
        ShowBox "Kick User", "Kick a user from NChat"
        If SelUser = "" Then Exit Sub
        Broadcast "kunø" & SelUser

    Case "N50UV"
        ShowBox "Get User Info", "Get a user's information"
        If SelUser = "" Then Exit Sub
        Broadcast "pipø" & SelUser

    Case "N500WX"
    'MsgBox "This item is unavailable. Please check back for info...", vbCritical, "Not available"
        Randomize
        Broadcast "fakeuø" & InputBox("Enter fake username") & "øDefaultø" & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255)

    Case "N50YZ"
        If TotalIcons < frmMain.ImageList1.ListImages.Count Then
            TotalIcons = TotalIcons + 1
        Else
            MsgBox "You have already purchased all the available icons! (" & frmMain.ImageList1.ListImages.Count & " in total)", vbCritical, "All Icons Purchased"
            Exit Sub
        End If

    Case "N500AA"
        MessageBold = True

    Case "N500BB"
        MessageUnderline = True

    Case "N500CC"
        frmMain.dlgSave.ShowColor
        MessageColour = frmMain.dlgSave.color

    Case "N500DD"
        frmMain.dlgSave.ShowColor
        MessageHColour = frmMain.dlgSave.color

    Case "N0DD"
        MessageBold = False
        MessageUnderline = False
        MessageColour = Msg
        MessageHColour = 0


    Case "N300EE"
        PlainText = InputBox("Enter the message everyone will see (Without clicking on it). For example: 'Click here to see answer to question'", "Spoiler Message", "Click here to read message")
        HiddenText = InputBox("Enter the message people will see when they click on the message. For example: 'This is the answer to the question'", "Spoiler Message", "This is the answer")

        PlainText = Replace(PlainText, Chr(34), "'")
        HiddenText = Replace(HiddenText, Chr(34), "'")

        If PlainText > "" And HiddenText > "" Then
            Broadcast "msgø<b onclick=" & Chr(34) & "alert('" & HiddenText & "')" & Chr(34) & ">" & PlainText & "</b>ø" & MessageColour & "ø" & MessageBold & "ø" & MessageUnderline
        Else
            MsgBox "You didn't enter one or more texts! No charge for this item", vbCritical, "Some messages not entered!"
            Exit Sub
        End If

    Case "N300FF"
        ToMarquee = InputBox("Enter text to scroll. The text will scroll from left to right on everyone's screen", "Send Marquee Text", "Ha Ha Ha... Dislocation")
        If ToMarquee > "" Then
            Broadcast "msgø<marquee>" & ToMarquee & "</marquee>ø" & MessageColour & "ø" & MessageBold & "ø" & MessageUnderline
        Else
            MsgBox "You didn't enter a message! Message was NOT sent, and you have not been charged for the item", vbCritical, "No message sent!"
            Exit Sub
        End If
    End Select

    NCredits = NCredits - Cost
    Text "Purchased " & tvItems.SelectedItem.Text & " for " & Cost & " NCredits with +ncredits+ NCredits remaining" & vbCrLf, Heading, True, , , 3
End Sub

Private Sub Form_Load()

    tvItems.Nodes.Add , , "H1", "Messages", ImageList1.ListImages(1).Key
    tvItems.Nodes.Add , , "H2", "NCredits", ImageList1.ListImages(1).Key
    tvItems.Nodes.Add , , "H3", "NChat Users", ImageList1.ListImages(1).Key
    tvItems.Nodes.Add , , "H4", "Icons / Colours", ImageList1.ListImages(1).Key
    tvItems.Nodes.Add , , "H5", "Cool Stuff", ImageList1.ListImages(1).Key


    For I = 1 To tvItems.Nodes.Count
        tvItems.Nodes.Item(I).Expanded = True
    Next I
' You can add your own items here. Here is the syntax:
'   N (MUST include, lets NChat know it's an actualy item
'   Price (Any length, MUST be numbers)
'   2 letter item code (MUST be 2 letters, or numbers, symbols ex. | if required)
'   The divider "|" without the quotes
' Finally,
'   The description of the item (When you click "Item Information")

' Then code the rest of it in as needed (ie. in cmdBuy_Click, using the same syntax as the other cases)
    tvItems.Nodes.Add "H1", 4, "N50AB|This item allows you to send a bold, purple message (often called a server message) to everyone in the current room", "Send a purple message", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H1", 4, "N150CD|This item allows you to send a fake message. It is sent using someone else's name, or a name that you choose", "Send a fake message", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H1", 4, "N100EF|When you send a message with this item, it will be sent with the name 'Anonymous User'", "Send an anonymous message", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H1", 4, "N200GH|This item sends a LARGE message. It is bigger than your normal message, and is useful for getting someone's attention", "Send a LARGE message", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H1", 4, "N150IJ|When you send a message with this item, it appears as an action, but with a different username that you can choose", "Send a fake action", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H1", 4, "N1000KL|This expensive item lets you send a purple message to EVERYONE in NChat, even if they are in another room. If you are looking for someone, then try sending this", "Send a purple message to EVERYONE", ImageList1.ListImages(2).Key

    tvItems.Nodes.Add "H2", 4, "N300MN|When you buy this item, you NChat will add or remove a random amount of ncredits from your total. This is really a game of chance", "Purchase random NCredits", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H2", 4, "N100OP|To play this game, pick a lucky number between 1 and 100. If your number is right, then you win 10,000 NCredits", "Lucky Numbers", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H2", 4, "N1000QR|This item will steal a random number of NCredits of someone. It will steal between 1 and 2000 NCredits off them", "Steal some NCredits off someone", ImageList1.ListImages(2).Key

    tvItems.Nodes.Add "H3", 4, "N500ST|This item will let you kick someone from NChat. They will exit NChat and have to re-connect", "Kick an NChat user", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H3", 4, "N50UV|When you use this item, you can view advanced information about a user, such as how many NCredits they have, as well as many other options", "View someone's information", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H3", 4, "N500WX|This item will let you add a fake user to the NChat chat room. This can be useful to fool someone into thinking that a friend has connected", "Create a fake user", ImageList1.ListImages(2).Key

    tvItems.Nodes.Add "H4", 4, "N50YZ|If someone has an NChat user icon that you don't have, then you can buy it from here.", "Purchase a new user Icon", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H4", 4, "N500AA|When you buy this item, every message sent from you will have your name in bold. It looks cool and draws attention to yourself :)", "Bold Username", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H4", 4, "N500BB|When you purchase this item, your messages will appear, with your name underlined. Just like the item above, but cooler :)", "Underline Username", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H4", 4, "N500CC|When you purchase this item, your username will appear in a different colour. Totally useless, but looks cool...", "Coloured Username", ImageList1.ListImages(2).Key
    'tvItems.Nodes.Add "H4", 4, "N500DD|When you purchase this item, your username will be highlighted in a really cool colour. Totally useless, but looks cool...", "Highlighted Username", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H4", 4, "N0DD|This item will reset your username colours, and their bold / underlined status. It is free of charge to buy", "Reset Username", ImageList1.ListImages(2).Key

    tvItems.Nodes.Add "H5", 4, "N300EE|This item will let you send a special message. When someone clicks on that message, it will display a box with a custom message inside it. Really usefull for information that not everyone may want to see, like the answer to a question, joke etc.", "Spoiler Message", ImageList1.ListImages(2).Key
    tvItems.Nodes.Add "H5", 4, "N300FF|Purchasing this item will let you send scrolling text to everyone. The text will scroll from left to right", "Send Marquee ", ImageList1.ListImages(2).Key
End Sub

Private Sub cmdInfo_Click()
    On Error GoTo NoShow

    Splice = SplitVB5(tvItems.SelectedItem.Key, "|")
    MsgBox Splice(1)
NoShow:
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    Me.Caption = "Welcome to the NChat Shop! " & NCredits & " NCredits remaining"
End Sub

Private Sub tvItems_Click()
    Splice = SplitVB5(tvItems.SelectedItem.Key, "|")
    Select Case Splice(0)
    Case "H1", "H2", "H3", "H4", "H5"
        Exit Sub
    Case Else
        Label1.Caption = "This item costs: " & Mid(Splice(0), 2, Len(Splice(0)) - 3) & " NCredits"
    End Select
End Sub

Private Sub tvItems_DblClick()
    cmdBuy_Click
End Sub
