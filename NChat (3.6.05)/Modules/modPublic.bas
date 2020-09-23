Attribute VB_Name = "modPublic"
' Any public, global or enum that doesn't fit into a category, ends up here
' Sure some of them can be categorised, but what are you going to do?
Option Compare Text
' StartMsg and EndMsg are what usernames are enclosed in
' For example: || Grayda || How are you man?
Public StartMSG As String
Public EndMSG As String

'
Public HTMLFile As String

' This is what user you picked from frmKick
Public SelUser As String

Public OldTickCount As Long
' Your 'encrypted' administrator password
Public Enc As String

' This string tells you or someone else who your
' last message was from. It is sent as a username,
' and not an IP address
Public LastFrom As String

Public UserName As String

' Allows you to have up to 3 Ini Files open
' Although only one is used in this code :)
Public IniFile(1 To 3) As String

' What files you have shared (*.mp3, *.jpg etc.)
Public FilePattern As String
' Share you files with the rest of those cheapos? ;)
Public ShareMyFiles As Boolean
' Your Automatic "Away Message"
Public AwayMessage As Boolean
' Whether or not 255.255.255.255 is used or 255.255.255.255
Public Loopback As Boolean
' Are you a true admin? (True admins cannot be kicked)
Public TrueAdmin As Boolean
' Allow swearing?
Public Swearing As Boolean
' Show smileys (The pictures, not just the code)?
Public Smiley As Boolean
' Whether or not to show the tip of the day-O
Public DontShowTip As Boolean



' How many NCredits you have
Public NCredits As Long
' How long you have been on NChat for
Public NChatTime As Long
' Your user icon that you have selected
Public MyIcon As String
' How many "icons" you actually own
Public TotalIcons As Integer



Public NewMessage As Boolean

' Allows us to access File System Commands
' Through the Scripting Library Reference
Public FileObj As New Scripting.FileSystemObject

' NChat version. 2 Digit Day, 2 Digit Month,
' 24 hour-hour, and minute
Public Const Ver = "NChat Build11 0903051359"

' This is our file version. This is NEW FILE VERSION ONE FOR BUILD 11
Public Const FileCheck = "NEWVER1"

' Our fancy message setup :)
Public MessageBold As Boolean
Public MessageUnderline As Boolean
Public MessageColour As String
Public MessageHColour As String

' This is our scripting class file, which basically points to the Broadcast and Text
' subs in their respective modules
Public scSubs As cSubs

' Our HTML Profile
Public Profile As String

' Should Notch be allowed to take advantage of scripting features?
Public AllowScripting As Boolean

' This public type, found in modZUser, lets us retrieve information about a user
' that sent the last message
Public RUser As RemoteUserDetails
