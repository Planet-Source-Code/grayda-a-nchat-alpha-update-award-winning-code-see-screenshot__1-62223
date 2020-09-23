Attribute VB_Name = "modMisc"
Option Compare Text
' Reads information from an INI or similarly structured
' text based file. Based off icey.gouranga.com's code
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Gets the proper windows username, instead of using environ("Username")
Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

' The API to detect how fast NChat Loads
' To use it, the first tick count is recorded
' (eg. 12345) and then later on the second count
' (eg. 34567) then the two are subtracted
' (34567 when started - 12345 when loaded = 22222 ms)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Example by Robin (rbnwares@edsamail.com.ph)
'Visit his site at http://members.fortunecity.com/rbnwares1
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Sub WriteString(SectionName As String, KeyName As String, KeyValue As String, INI As Integer)
'KPD-Team 1999
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net

    WritePrivateProfileString SectionName, KeyName, KeyValue, IniFile(INI)
End Sub


Public Sub WriteSect(SectionName As String, DefaultKey As String, INI As Integer)
    Call WritePrivateProfileSection(SectionName, DefaultKey, IniFile(INI))
End Sub



Public Function ReadText(Sec As String, Key As String, Num As Integer)
' Allows you to read items from an INI file
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadText = Left(sRet, GetPrivateProfileString(Sec, ByVal Key, "", sRet, 255, IniFile(Num)))
End Function


Sub Main()
'On Error Resume Next
' Main routine, which sets up profiles
' and other information

' Starts our 'timer'. When this is called
' again, you can subtract the old value from
' the new value, to determine how long has elapsed
    OldTickCount = GetTickCount

    ' Sets up our 'default' colours, or else everything will be black ( :O )
    Msg = 16738603
    'Msg = 0
    con = 32768
    dis = 202
    svr = 8388736
    act = 33023
    ThatsGood = 49408
    ThatsBad = vbRed
    Heading = 12615680

    MessageColour = Msg

    ' Activate smileys. You can turn it off
    ' you if get annoyed at the "cute" pictures.
    ' Most smileys are from Jayanth Kumar J's
    ' Network Chat, also available from
    ' planetsourcecode.com. :hutchy, :lick, :idge
    ' :big gay al smileys created by me,
    ' no offence intended of course
    Swearing = True
    Smiley = True

    ' our address that our winsocks will run off.
    ' you can change this by choosing loopback
    ' from the admin menu
    Address = "255.255.255.255"

    'On Error Resume Next

    ' Copy our splash screen to the app's folder
    CopyFromRes 101, "PNG", "test.png"

    ' "User Icons" folder doesn't exist? Make it then, for storing custom icons and stuff
    'If FileObj.FolderExists(AppPath & "User Icons") = False Then MkDir AppPath & "User Icons"

    'frmMisc.File1.Path = AppPath & "User Icons"
    'For i = 1 To frmMisc.File1.ListCount - 1
    'frmMain.ImageList1.ListImages.Add frmMain.ImageList1.ListImages.count + 1, frmMisc.File1.List(i), stdole.LoadPicture(AppPath & "User Icons\" & frmMisc.File1.List(i), 32, 32)
    'Next i

    frmMain.Show

End Sub

Public Function Decode(ByVal iString As String, iKEY As String) As String
' Decrypts strings and integers
' Used in the save file etc.
    Dim Password As String
    Dim Words As String
    Dim Encrypted As String
    Dim Tempchar As String
    Dim Tempchar1 As String
    Dim Counter As Integer
    Dim TempAsc As Integer
    Dim TempAsc1 As Integer
    Counter = 1
    Password = iKEY
    Words = iString


    For x = 1 To Len(Words)    'loop for Each letter of the password

        Tempchar1 = Mid(Password, Counter, 1)    'get a Single letter of the password
        Tempchar = Mid(Words, x, 1)    'get a Single letter of the words

        TempAsc = Asc(Tempchar)
        TempAsc1 = Asc(Tempchar1)
        TempAsc = TempAsc - TempAsc1

        If TempAsc < 0 Then TempAsc = TempAsc + 245

        Tempchar = Chr(TempAsc)
        Encrypted = Encrypted & Tempchar
        Counter = Counter + 1    'incriment the counter

        If Counter > Len(Password) Then Counter = 1

    Next x
    Decode = Encrypted

End Function


Public Function Encode(ByVal iString As String, iKEY As String) As String
' Encrypts strings and integers
' Using a password (key)
    Dim Password As String
    Dim Words As String
    Dim Encrypted As String
    Dim Counter As Integer
    Dim Tempchar As String
    Counter = 1
    Password = iKEY
    Words = iString


    For x = 1 To Len(Words)    'loop for Each letter of the password

        Tempchar1 = Mid(Password, Counter, 1)    'get a Single letter of the password
        Tempchar = Mid(Words, x, 1)    'get a Single letter of the words

        TempAsc = Asc(Tempchar)    'convert the letter of the password To a number
        TempAsc1 = Asc(Tempchar1)    'convert the letter of the word To a number
        TempAsc = TempAsc + TempAsc1    ' add the two values

        'check to see if the value if greater than 245. if it is,
        'subtract 245 from it.
        'this makes sure that we don't go past the highest ascii value

        If TempAsc > 245 Then TempAsc = TempAsc - 245

        Tempchar = Chr(TempAsc)    'convert the number back To a character

        Encrypted = Encrypted & Tempchar    'add the character To the End of the encrypted String
        Counter = Counter + 1    'incriment the counter

        'check to see if the counter is > the
        '     length of the password
        'if it is, set the counter to 1
        If Counter > Len(Password) Then Counter = 1

    Next x
    'show the encoded text in the textbox
    Encode = Encrypted

End Function



Sub CopyFromRes(ByVal ID As Integer, ResType As String, FileName As String, Optional Path As String)
' Copies the files from the resource file into
' the current directory so they can be called faster
    tosave = AppPath & Path
    tosave = tosave & FileName

    Dim A As Long

    Open tosave For Output As #4
    Print #4, StrConv(LoadResData(ID, ResType), vbUnicode);
    Close #4

End Sub

Public Function GetUserName() As String
' Simple sub to get our windows username
    Dim UserName2 As String * 255
    Call GetUserNameA(UserName2, 255)
    GetUserName = Left$(UserName2, InStr(UserName2, Chr$(0)) - 1)
End Function

Public Function AppPath() As String
    If Right(App.Path, 1) = "\" Then
        AppPath = App.Path
    Else
        AppPath = App.Path & "\"
    End If
End Function

' This is the NEW sub which can extract elements from a HTML file.
' It does this by finding the element in question (ie. "color:", then
' getting the text between "color:" and ";", which essentially is our CSS element
Public Function GetElement(Element As String, TagToGet As String, HTMLFile As String)
    GetElement = ExtractFromTags(ExtractFromTags(HTMLFile, "." & Element & " {", "}"), TagToGet & ": ", ";")
End Function

Public Sub RunHyper(Hyperlink As String)
' Run our hyperlink using an extended shell command
    lngRet = ShellExecute(0&, "Open", Hyperlink, "", vbNullString, 1)
End Sub
