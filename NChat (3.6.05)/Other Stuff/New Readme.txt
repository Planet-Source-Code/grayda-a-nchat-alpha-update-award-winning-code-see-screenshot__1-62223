---------------------------
NChat Alpha Build 10 Readme
---------------------------

This readme file contains notes, bugs, last minute additions and other stuff relating to the NChat Source Code and it's use, plus technical information. It is NOT a basic how-to-use guide, because NChat already has it's own help file, available in the "Help" folder within this directory


Section 1: What IS NChat Alpha?

NChat Alpha (Or just NChat) functions as both a SERVERLESS Network Chat program, for use on a local area network (or LAN), AND as a chat program for use over the internet (or Network if needed), using a server, which is included in this folder. It has a whole host of features that you simply won't find in any other chat program for both the internet and network such as:

> Now Internet compatible!
> Over 60 Picture Smileys, ready to insert
> Customizable pictures next to your name
> FULL set of Administrator tools including:
   > Kick users
   > Re-Direct Users
   > Ban users by IP or username
   > Ghost users
   > Print User Info
> A fully working chat-bot (Called Notch, the NChat AutoBot) with:
   > Wildcard searching and matching
   > INI Based structure for easy programming
   > Multiple answers to a question ensure bot is never monotonous
   > Bot has idle chatter, determined after a certain interval
   > Bot can be updated with new content on-the-fly!
> Earn NCredits - NChat's very own currency!
> Spend NCredits in the store, and purchase some cool items!
> Full data logging, for people who want to learn more about NChat's data system
> WinMX Style chatroom "actions"
> Custom, funky green command buttons
> Load colour profiles to change NChat's appearance
> Private chat with unlimited people, with private Whiteboard!
> NChat supports Unlimited users, only limited by your bandwidth!
> Create and run your own chat room!
> Awesome Transparent PNG Splash screen
> Almost totally compatible with Windows 98 and ME!

> Plus HEAPS more features

When using NChat as a stand-alone Network Chat, you don't need any servers AT ALL! Each copy of NChat is both a client AND a server! You can host your own room without needing any extra software.

When using NChat as an internet chat program, you need to run the NChat server, available in this folder. You simply run it on a computer with a static IP (basically, any computer that CAN act as a web-server) and a reliable internet connection, and give your IP address to potential connecters.


Section 2: What software does NChat require?

NChat requires the following files:

msvbvm60.dll <-- Standard Visual Basic File
oleaut32.dll <-- OLE Automation Library
olepro32.dll
asycfilt.dll 
stdole2.tlb  <-- OLE Automation
COMCAT.DLL 
msimg32.dll  <-- For Alpha Blending and other imaging stuff
GdiPlus.dll  <-- For Windows 2000 / XP ONLY. More imaging stuff for splash screen
scrrun.dll   <-- For File System Objects and other scripting stuff
RICHTX32.OCX <-- Standard Rich Text Box. A beefed up Text Box
comdlg32.ocx <-- Common Dialog control. Shows Open, Save, Print and colour boxes
MSCOMCTL.OCX <-- Common Windows Controls such as Progress Bars, Image Lists etc.
TABCTL32.OCX <-- Tabbed Dialog. Lets you group controls in a tabbed-interface
MSWINSCK.OCX <-- The heart of NChat. Lets us use Network Resources such as UDP and TCP

If you do not have these files, then you can download the NChat Alpha setup file from:

http://www.solidinc.tk (Under the Downloads page, under the Applications category)

The setup file will install NChat Alpha Build 10, plus the required dependency files, or you can search on Google for the file names, or better yet, upgrade to Windows XP with VB6.0!

NChat has been sucessfully tested on the following Operating Systems:

Windows XP 	(100%, Developed mostly on this OS)
Windows 2000 	(99% The older versions were developed on this)
Windows ME	(80% Using VMWare computer emulation)
Windows 98	(Not yet tested, CD's busted)
Windows 95	(Not yet tested, CD's busted too)
Linux with Wine (0%, but tested with old version of NChat and Wine)

This build of NChat was tested with the following computer specifications:

Pentium 4 1.5 GHZ with 128MB RAM, Windows XP Home
Pentium 1, 166 MHZ with 64mb RAM, Windows XP Home

Average Load time for P4: 1511ms, when run straight from IDE
Average Load time for P1: 6509ms, When run straight from IDE

If you have tested NChat on any of the operating systems which haven't been tested (or even emulated OS's), then please forward the results to: firestorm_visual@hotmail.com to be included in the about box in the next major release. (Big accolades, I know :P)


Section 3: Current Problems with NChat Alpha Build 10

NChat has some major issues that make it inappropriate for use in high-risk areas where danger is commonplace, such as on oil-rigs, places dealing with explosives, or high levels of toxic chemicals. If delivery of messages is CRITICAL (Really), then consider purchasing a TCP/IP chat program, or use MSN messenger, which is developed by Microsoft, a reputable (:stifled laugh) company. Current problems are as follows:

> The server portion of this software (using TCP/IP Mode) has NOT been tested over internet, or network, and is NOT guaranteed to work

> File transfer is NOT available at this time because of control difficulties. This is currently being worked on

> ALL messages to and from NChat are NOT secure. They are not encrypted in ANY way so anyone who is an NChat Admin, or has a packet sniffer can read your conversations. This includes Private Messages (PMs)

> Some messages are NOT delivered when using NChat in UDP mode (Client / Server combination). This is a known weakness in UDP technology, and there is NO known workaround.

> Some messages are delivered -twice- due to UDP technology. No known workaround yet (Except for a dirty little hack, but I don't wanna use that :) )

> Notch is stupid. He doesn't have contextual functions yet, so if notch is looking for: "Notch is cool" then he will see "Notch is NOT cool". There MAY be a fix added later, but don't count on it!

> There may be some other high-risk bugs that I haven't figured out yet. Maybe someday


Section 4: About Notch

Notch is NChat's room robot. Also known as a chatbot, it's function is to enter the chat room, and let people talk to it, just like a real person. The only difference is that Notch can be programmed. He can say what you tell him to say, do what you tell him to do and so forth. He is NOT AI, rather like a parrot. You say something, he says something back. No intelligence yet. He can learn new phrases, but will only learn what you say, not what he thinks is the best response.

All that aside, Notch uses a structured INI file, looking like this:

[Phrase1]
Question=This is the question
Answer1=This is answer #1
Answer2=This is answer #2

etc.

Notch has the ability to understand PARTS of sentences. So if you type Hi Notch, how are you man?, he can understand Hi, Notch and Man (Which he has in his list of words, or DATABASE)

Notch has the potential to be better, with contextual stuff, and more intelligent behaviour, but because this is an OPEN chat, with heaps of people chatting at once, he can't be expanded as I feel is necessary.


Section 5: The help file

Yup. NChat now has a fully loaded .hlp file. It's pretty basic at the moment, and has taken me about a month on and off to develop. It contains advanced details for administrators, Notch help integrated into it, plus some other stuff for hosting NChat over the internet. Because of size issues, the help file hasn't been included with the source code, or the pre-compiled binary release available on www.solidinc.tk. Instead, you must download it from www.solidinc.tk and it's about 1-2mb big I think. :)