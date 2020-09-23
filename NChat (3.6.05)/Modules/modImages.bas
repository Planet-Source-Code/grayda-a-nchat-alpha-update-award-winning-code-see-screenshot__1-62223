Attribute VB_Name = "modImages"
' This is the NEW modImages.
' Coz converting Images to RTF codes is no longer necessary, this section
' has been GREATLY reduced. It's now mainly used for the cool splash screen
' that's loaded in frmSplash
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Option Compare Text


' Stuff for making things transparent :)
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

Public ThePic As String
Public PicStarted As Boolean

' This is for both the Per-Pixel PNG rendering,
' and
Public Declare Function SetWindowLong Lib "user32" _
                                      Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                              ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Our API to flood fill a picture box
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

' Just some simple Publics for the whiteboard
' They tell you where your last line was drawn from
Public lastX As Long
Public lastY As Long

Private Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Private Type METAHEADER
'    mtType As Integer
'    mtHeaderSize As Integer
'    mtVersion As Integer
'    mtSize As Long
'    mtNoObjects As Integer
'    mtMaxRecord As Long
'    mtNoParameters As Integer
'End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, bitmap As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, bitmap As Long) As GpStatus

Public Type GdiplusStartupInput
    GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
    DebugEventCallback As Long          ' Ignored on free builds
    SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
    ' the hook/unhook functions properly
    SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
    ' its internal image codecs.
End Type


Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Public Enum GpStatus   ' aka Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum

Public Function getTempName(Optional anExt As String = "tmp") As String
' This retrieves the temp path on your drive
' eg. c:\windows\temp
    Dim tempPath As String
    Dim FileName As String
    Dim I As Long

    Const validChars As String = "123567890qwertyuiopasdfghjklzxcvbnm"

    ' Create a buffer
    tempPath = String$(255, " ")
    ' get the system path
    GetTempPath 255, tempPath
    ' trim off the fat
    tempPath = Left$(tempPath, InStr(tempPath, Chr$(0)) - 1)
    ' Create a buffer
    FileName = Space(12)
    ' Put the non-random stuff into the string
    Mid$(FileName, 1, 1) = "T"
    Mid$(FileName, Len(FileName) - Len(anExt), 1) = "."
    ' Add in the specified extension, if provided ("tmp" is default)
    Mid$(FileName, Len(FileName) - Len(anExt) + 1, Len(anExt)) = anExt
    ' fill the buffer with random stuff
    Randomize
    For I = 2 To Len(FileName) - 4
        Mid$(FileName, I, 1) = Mid$(validChars, CLng(Rnd() * (Len(validChars)) + 1), 1)
    Next I
    tempPath = tempPath & FileName
    ' return the path name
    getTempName = tempPath

End Function
