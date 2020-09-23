VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub SendStuff(Data As String)
' This is simple. When people call SendStuff from the script, It just calls
' the broadcast sub. If we had of called Broadcast directly (ie from modWSock),
' then they could mess with the other functions and subs in modWSock, such as
' binary arrays and stuff, and could cause NChat to crash
    Broadcast Data
End Sub

Sub WriteText(TheText As String, Optional Colour As String, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As String, Optional Font As String, Optional CheckSmileys As String)
' lol same as above. All about security etc.
' Doesn't want to handle colour shortcuts yet (ie. msg, con, thatsbad, heading etc.)
    Text TheText, Colour, Bold, Italic, Underline, Size, Alignment, Font, CheckSmileys
End Sub

Function Data() As String
' Lets our script handle some sData stuff
    Data = sData
End Function

Function sResult(Number As Integer)
' Let our script-ees see Results just like NChat can
    sResult = Result(Number)
End Function

Function Wait()
    DoEvents
    DoEvents
End Function

Function sLastFrom() As String
    sLastFrom = LastFrom
End Function
