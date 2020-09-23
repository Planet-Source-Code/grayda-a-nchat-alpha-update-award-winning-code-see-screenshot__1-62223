Attribute VB_Name = "modScripting"
' This mod is for NChat administrators to extend Notch's power. He can now handle
' SIMPLE scripts, which can let him do stuff like read from files etc.

Public Sub PrepareScripting()
' This sub prepares scripting for use. It simply adds "Broadcast" and "Text"
' to our scripting control. For security reasons, the other functions like
' NCredits and stuff are off-limits.

' Set up our scripting control for use with Notch

' Add our cSubs to the control. This lets us use broadcast and text etc.
    Set scSubs = New cSubs
    frmMain.scNotch.AddObject "SCRIPT", frmMain.scNotch, True
    frmMain.scNotch.AddObject "SUBS", scSubs

End Sub

Public Sub RunScript(ScriptFile As String)
' If the scripts is in out "<Location of NChat>\Other Stuff\Scripts\" folder, then
' you don't need to type "C:\Blah\Test.vbs", just "Test.vbs"

' Otherwise, just run it...
    If Mid(ScriptFile, 2, 2) <> ":\" Then
        ScriptFile = AppPath & "Other Stuff\Scripts\" & ScriptFile
    End If

    If FileObj.FileExists(ScriptFile) = False Then
        MsgBox "Notch's scripting commands have encountered an error: " & ScriptFile & " doesn't seem to exist. Please check the file exists, or remove the reference from " & frmAutobot.Text1.Text & ". Execution of script will now halt", vbCritical, "File doesn't exist!"
        Exit Sub
    End If

    ' Open the file, and directly run what's in our file
    Open ScriptFile For Input As #1
    frmMain.scNotch.ExecuteStatement Input(LOF(1), 1)
    Close #1
End Sub
