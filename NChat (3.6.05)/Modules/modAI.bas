Attribute VB_Name = "modAI"
Option Compare Text

Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Is Notch Running?
Global NotchRunning As Boolean
' Is Notch Learning?
Public NotchLearning As Boolean
' Who he is learning from
Public LearningFrom As String
' and the word he is learning
Public LearningWord As String

' How many new words Notch has learnt
Global NewWords As Integer

' EnumSect lets us enumerate our INI section names
' and in return get the number of sections, so Notch
' can work out how many questions he has
Public Function EnumSect() As Integer
    Dim szBuf As String, Length As Integer
    Dim SectionArr() As String
    szBuf = String$(255, 0)
    Length = GetPrivateProfileSectionNames(szBuf, 255, IniFile(2))
    szBuf = Left$(szBuf, Length)
    SectionArr = Split(szBuf, vbNullChar)
    EnumSect = UBound(SectionArr)
End Function

Public Sub DoAutoBot(QuestionToAutoBot As String)

' Our question from the INI file, split up into smaller bits
    Dim MultiQ() As String
    Dim ToCheck As String
    Dim ThePhrase As String
    Dim R As Integer

    If NotchRunning = False Or LastFrom = "Notch" Then Exit Sub
    ' Loop thru all phrases except for [idle phrase] and [general]
    For I = 1 To EnumSect - 2

        ToCheck = ReadText("Phrase" & I, "Question", 2)

        MultiQ = SplitVB5(ToCheck, "||")

        For n = LBound(MultiQ) To UBound(MultiQ)

            ' Found our question? then pick a rando question then go for it!
            If WyldCard(QuestionToAutoBot, MultiQ(n)) = True Then

                ' Enumanswers lets us work out how many 'Answer#' keys we have in a phrase
                Randomize
                temp = EnumAnswers("Phrase" & I)

TryAgain:

                If temp > 1 Then
                    R = Int(Rnd * temp)
                Else
                    R = 1
                End If


                If R = 0 And temp = 0 Then GoTo TryAgain


                ThePhrase = ReadText("Phrase" & I, "Answer" & R, 2)

                ' If it's a command to run a script file (eg. vbs), then don't broadcast
                ' the phrase, instead let RunScript handle it, then exit the sub to avoid
                ' text like "%script=blahblahblah.vbs%" from showing up.
                ' Check out RunScript in modScripting for info
                If Left(ThePhrase, 8) = "%script=" And AllowScripting = True Then
                    RunScript Mid(ThePhrase, 9, Len(ThePhrase) - 9)
                    Exit Sub
                ElseIf Left(ThePhrase, 8) = "%script=" And AllowScripting = False Then
                    Exit Sub
                End If

                ThePhrase = Replace(ThePhrase, "%rss%", GetRSSHeadline(frmAutoBotOptions.Text1.Text))
                Broadcast ThePhrase
                DoEvents
                DoEvents
                ' Need to add MSG to PM check
            End If
        Next n
    Next I
End Sub

Public Function EnumAnswers(Section As String) As Long
    EnumAnswers = 0

    Do Until ReadText(Section, "Answer" & EnumAnswers + 1, 2) = ""
        EnumAnswers = EnumAnswers + 1
    Loop

End Function

Public Function Try2Learn()

' OK This is Notch's learning center. If someone sends a message to Notch, (msg or pm1)
' Notch is learning (NotchLearning = True), and he doesn't have the question
' in his DB, then he will learn it. He asks for a response, then when
' one is provided, he will let them know, write it into the DB, and reload
' the words to be used. Kinda simple... :|

' BTW, when Notch is learning. Only one person can teach at a time
' ie if Grayda is trying to teach Notch something, and he accidentally
' says 'Hi' to someone else, then Notch will interpret that as the response
' to his learning question. (Eg. Hi Notch. Response is: Hello dude!)

' Notch can't learn from himself, it would create
' a never ending loop!! (?)

    If LastFrom = "Notch" Then Exit Function
    If LearningWord = Result(3) Then Exit Function

    ' If we are currently learning something, then don't learn something else
    If Trim(LearningFrom) = "" And WyldCard(Result(3), "Notch") = True Then
        LearningWord = Result(3)
        LearningFrom = LastFrom

        WriteSect "Phrase" & EnumSect - 1, "Question=" & Result(3), 2
        If Result(1) = "msg" Then
            Broadcast "msgøNotchø+lastfrom+, I don't understand your question. Please provide me with a response!ø0øFalseøFalseø0"
            ' Double DoEvents so our message has time to reach the recipient before
            ' trying to do a "DoAutoBot" again
            DoEvents
            DoEvents

        ElseIf Result(1) = "pm1" Then
            Broadcast "pm1ø+lastfrom+ø+lastfrom+, I don't understand your question. Please provide me with a response!øNotch"
            ' Double DoEvents so our message has time to reach the recipient before
            ' trying to do a "DoAutoBot" again
            DoEvents
            DoEvents

        End If
        ' Wait until we get a response
        Exit Function
        ' Not learning anything (LearningFrom is blank)? then start learning!
    ElseIf LearningFrom <> LastFrom Then

        Exit Function
    ElseIf LastFrom = LearningFrom Then
        TotPhrases = EnumSect - 2


        If Result(1) = "msg" Then
            Broadcast "msgøNotchøThanks +lastfrom+! I now know what you are talking about!ø0øFalseøFalseø0"
            ' Double DoEvents so our message has time to reach the recipient before
            ' trying to do a "DoAutoBot" again
            DoEvents
            DoEvents

            WriteString "Phrase" & TotPhrases, "Answer1", "msgøNotchø" & Result(3) & "ø0øFalseøFalseø0", 2
        ElseIf Result(1) = "pm1" Then
            Broadcast "pm1ø+lastfrom+øThanks +lastfrom+! I now know what you are talking about!øNotch"
            ' Double DoEvents so our message has time to reach the recipient before
            ' trying to do a "DoAutoBot" again
            DoEvents
            DoEvents

            WriteString "Phrase" & TotPhrases, "Answer1", "pm1ø+lastfrom+ø" & Result(3) & "øNotch", 2
        End If

        NewWords = NewWords + 1
        LearningFrom = ""
        ' Reload our Notch, will NEW words included!
        OldINI = IniFile(2)
        IniFile(2) = ""
        IniFile(2) = OldINI
        Exit Function
    End If


End Function

' This is my custom WildCard system.
' It lets you use *'s to search for items
' Here is how it works:

' The Search Criteria (EG: Testing*Hello)
'   is SplitVB5 into an array.
' The first part (EG: Testing) is searched for
'   using the InStr command. If it is found, then
'   1 is added to the number of correct matches.
'   If it isn't found, nothing changes
' At the end, if the Number of matches is equal
'   to the number of items in the array, then
'   it returns as true, if not, then false

' If you can improve on this, then please let me know,
' by sending an e-mail to: firestorm_visual@hotmail.com
' but this fits my needs perfectly, so I don't think
' I'll improve on it :)
Public Function WyldCard(StringToCheck As String, SearchFor As String) As Boolean
    On Error Resume Next

    Dim Search() As String
    Dim Matches As Integer
    SearchFor = SearchFor & "*"
    Search = SplitVB5(SearchFor, "*")

    For n = LBound(Search) To UBound(Search)

        RetPos = InStr(1, StringToCheck, Search(n), vbTextCompare)

        If RetPos > 0 Then
            Matches = Matches + 1
        End If

    Next n

    If Matches = UBound(Search) + 1 Then
        WyldCard = True
    Else
        WyldCard = False
    End If


End Function



Public Function GetRSSHeadline(RSSURL As String) As String
' The easy way to extract a random <description></description> tag from
' an RSS-Feed. Tested with Yahoo!'s 'Oddly Enough' feed from
' Sun, 26 Dec 2004 02:05:14 GMT (http://news.yahoo.com/news?tmpl=index&cid=757)
    Dim XML As String
    Dim StartAt As Long
    If RSSURL = "" Then
        GetRSSHeadline = "Has someone got a copy of todays newspaper? I can't find mine :("
        Exit Function
    End If

    Open RSSURL For Input As #1
    XML = Input$(LOF(1), 1)
    Close #1
    Randomize
    StartAt = Int(Rnd * Len(XML))

    GetRSSHeadline = ExtractFromTags(XML, "<description>", "</description>", StartAt)
    ' Clean up the tags a bit
    GetRSSHeadline = Replace(GetRSSHeadline, "&#039;", "'")
    GetRSSHeadline = Replace(GetRSSHeadline, "&#151;", ":")
    GetRSSHeadline = Replace(GetRSSHeadline, "&#36;;", "$")
    GetRSSHeadline = Replace(GetRSSHeadline, "&quot;", Chr(34))    ' Quotation Marks "
    GetRSSHeadline = Replace(GetRSSHeadline, vbLf, " ")

    ' Change this to get a more 'human' response
    If Trim(GetRSSHeadline) = "" Then GetRSSHeadline = "Oh wait... sorry +lastfrom+, but my newspaper is blank. Turns out there is no news today! :)"


End Function

Public Function ExtractFromTags(Source As String, OpenTag As String, CloseTag As String, Optional StartFrom As Long) As String
' Lets us extract text from within HTML or XML document.
' Doesn't seem to work with tags set out like this:

' <a tag>
' Testing
' </a tag>

' ONLY Like this:
' <a tag>Testing</a tag>
    On Error Resume Next
    Dim StartPos As Integer
    Dim FinishPos As Integer

    If StartFrom = 0 Then StartFrom = 1

    StartPos = InStr(StartFrom, Source, OpenTag, vbTextCompare)
    FinishPos = InStr(StartPos + Len(OpenTag), Source, CloseTag, vbTextCompare)
    If StartPos <= 0 Then Exit Function
    'LastAt = FinishPos + 1
    StartPos = StartPos + Len(OpenTag)
    OldStartPos = StartPos
    FinishPos = FinishPos

    ExtractFromTags = Mid(Source, StartPos, FinishPos - StartPos)

End Function
