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
SUBS.SendStuff "msg�Notch�Who's up for a game of hangman? I am. I'm thinking of a word. It's " & len(CurWord) & " letters long!�0�False�0"

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
SUBS.SendStuff "msg�Notch�:wow You got it! The word was " & curword & "�0�False�0"
Else
SUBS.SendStuff "msg�Notch�Nope. Keep trying +lastfrom+�0�False�0"
Guesses = Guesses - 1

End If
If Guesses = 0 then 
SUBS.SendStuff "msg�Notch�Game over! The word I was looking for is: " & curword & "�0�False�0" 
Else
'Goto DoHang
End If


