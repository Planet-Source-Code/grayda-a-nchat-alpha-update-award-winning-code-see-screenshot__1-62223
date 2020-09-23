Attribute VB_Name = "modRTB"
Public Const DGreen = 32768
Public Const Orange = 33023
Public Const DRed = 196
Public Const Purple = 8388736
Public Sub Text(Text As String, txtObject As RichTextBox, Optional Colour As ColorConstants, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As AlignmentConstants, Optional Font As String)
On Error Resume Next
' Puts text into the rich text box
If Text = "" Then Exit Sub

With txtObject
' Set the cursor at the end
    .SelStart = Len(.Text)
' The length of the selection should be 0
    .SelLength = Len(.Text)
    
    .SelBold = Bold
    If Font > "" Then .SelFontName = Font
    .SelItalic = Italic
    .SelUnderline = Underline
    .SelFontSize = Size
    .SelAlignment = Alignment
    .SelColor = Colour
    
    .SelText = Text
    .SelStart = Len(.Text)
    .SelLength = 0

End With
End Sub
