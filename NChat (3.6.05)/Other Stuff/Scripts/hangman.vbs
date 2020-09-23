' OK this is a half-arsed attempt at a hangman game. It uses some loops and other crap.
' uhhhh. I'll update the comments later

' Our word being guessed
Dim CurWord
' How many guesses remaining.
Dim Guesses
' Our old data
Dim OldData

Dim Part1
Dim Part2

Guesses = 5
CurWord = "TestWord"

' Tell everyone we're playing hangman
SUBS.SendStuff "msgøNotchøWho's up for a game of hangman? I am. I'm thinking of a word. It's " & len(CurWord) & " letters long!ø0øFalseø0"

OldData = SUBS.Data
'Goto DoHang

Do until OldData <> SUBS.Data
SUBS.Wait
Loop

If SUBS.sLastFrom = "Notch" then 

OldData = SUBS.Data
Part1 = Instr(1,SUBS.Data,"HANGMAN", vbTextCompare)
Part2 = Instr(Part1,SUBS.Data," ", vbTextCompare)
If Part1 > 0 then Word = mid(SUBS.Data, Part1, Part2 - Part1)

If CurWord = Word then
SUBS.SendStuff "msgøNotchø:wow You got it! The word was " & curword & "ø0øFalseø0"
Else
SUBS.SendStuff "msgøNotchøNope. Keep trying +lastfrom+ø0øFalseø0"
Guesses = Guesses - 1

End If
If Guesses = 0 then 
SUBS.SendStuff "msgøNotchøGame over! The word I was looking for is: " & curword & "ø0øFalseø0" 
Else
'Goto DoHang
End If


