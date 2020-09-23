' This is the example script file for Notch, written by Grayda of Solid Inc. Media Productions:
' www.solidinc.tk or www.nchat.tk

' To get more scripts like this, then try the NChat web-site: www.nchat.tk

' This script will show scripts in action. If you want to learn how to write your own NChat-compliant
' scripts, then check out the help file, included with both the source code and the installer

' First, write some text:

SUBS.WriteText "Starting example script: ",vbBlack,true
SUBS.WriteText "DONE!" & vbcrlf,vbBlue,true
SUBS.WriteText "Sending example message to everyone in the room: " & vbcrlf,vbBlack,true
SUBS.WriteText "DONE!" & vbcrlf,vbBlue,True

' Then send the aforesaid message to everyone. 

SUBS.SendStuff "msgøNotchøI'm reading this line from a script. Gee I'm smart :)ø0øFalseø0"

' Return the command that was last sent

msgbox "The last sample of Data received by NChat was: " & SUBS.Data,vbExclamation,"Example of SUBS.Data"
' It's done!

SUBS.WriteText "This is the end of the script. Closing..." & vbcrlf,vbBlack,true

' And that's it! For more info, read the help file included with the sourcecode / installer

' To put this script to use, Open your Notch INI file and add a phrase like this:

' [Phrase1]
' Question=Can you script?
' Answer1=%script=example.vbs%

' If example.vbs is in <Location of NChat>\Other Stuff\Scripts, then just type example.vbs, BUT if it's in
' c:\TestFolder\example.vbs, then type the FULL path and filename.