PLEASE NOTE. THIS BUILD OF NCHAT ALPHA WILL MOST LIKELY **NOT** BE UPDATED. THIS VERSION IS THE FINAL VERSION, HOWEVER UPDATES MAY CHANGE. FOR THE **NEW** NCHAT, VISIT THE WEB-SITE WWW.SOLIDINC.TK. THE LINK TO WWW.NCHAT.TK IS **NOT** VALID, AS ITS OWNED BY SOMEONE ELSE. ERMM... ENJOY THESE CAPS :P

---------------------------
NChat Alpha Build 12 Readme
---------------------------

Please maximize this window to view the document in it's originally intended form

This file has been included to answer some commonly asked questions about NChat. It's not a basic how-to, but rather a technical document. It answers questions such as: "Why doesn't NChat work?", "How can I understand your source code?" and some others. For more information, send me an e-mail at: firestorm_visual@hotmail.com or visit the forums at www.solidinc.tk

---------------------------
Section 1: What is NChat?
---------------------------

NChat is a network / internet chat application. You can use it on your local network without a server, or use it over the internet with a specially designed server. NChat also offers features you won't find on any network / internet chat program:

> Over 40 Picture Smileys, ready to insert
> Customizable pictures next to your name
> FULL set of Administrator tools including:
   > Kick users
   > Re-Direct Users
   > Ban users by IP or username
   > Ghost users
   > Print User Info
   > Add and remove fake users
> A fully working chat-bot (Called Notch, the NChat AutoBot) with:
   > Wildcard searching and matching
   > INI Based structure for easy programming
   > Multiple answers to a question ensure bot is never monotonous
   > Bot has idle chatter, determined after a certain interval
   > Bot can be updated with new content on-the-fly!
   > Bot can keep up-to-date with the news using RSS feeds
> Earn NCredits - NChat's very own currency!
> Spend NCredits in the store, and purchase some cool items!
> Full data logging, for people who want to learn more about NChat's data system
> WinMX Style chatroom "actions"
> Load colour profiles to change NChat's appearance
> Private chat with unlimited people, with private Whiteboard!
> NChat supports Unlimited users, only limited by your bandwidth!
> Create and run your own chat room!
> Awesome Transparent PNG Splash screen
> Almost totally compatible with Windows 98 and ME!

---------------------------
Section 2: What software does NChat require?
---------------------------

NChat requires the following files:

msvbvm60.dll <-- Standard Visual Basic File
oleaut32.dll <-- OLE Automation Library
olepro32.dll <-- More OLE stuff
asycfilt.dll <-- No description available
stdole2.tlb  <-- OLE Automation
COMCAT.DLL   <-- Microsoft Component Category Manager Library
msimg32.dll  <-- For Alpha Blending and other imaging stuff
GdiPlus.dll  <-- For Windows 2000 / XP ONLY. More imaging stuff for splash screen
scrrun.dll   <-- For File System Objects and other scripting stuff
RICHTX32.OCX <-- Standard Rich Text Box. A beefed up Text Box
comdlg32.ocx <-- Common Dialog control. Shows Open, Save, Print and colour boxes
MSCOMCTL.OCX <-- Common Windows Controls such as Progress Bars, Image Lists etc.
TABCTL32.OCX <-- Tab control
MSWINSCK.OCX <-- The heart of NChat. Lets us use Network Resources


If you don't have these files, then you can download the NChat setup file from: www.solidinc.tk under "Downloads"

NChat has been sucessfully tested on the following Operating Systems:

Windows XP 	(100%)
Windows 2000 	(99%)
Windows ME	(80%)
Windows 98	(Not yet tested)
Windows 95	(Not yet tested)
Linux with Wine (0%, but tested with old version)

---------------------------
Section 3: Why doesn't NChat work?
---------------------------

This is a VERY general question, but we should narrow it down a bit.

Q: Why can't I see any names (including mine) on the user list?
A: For NChat to work, you need a valid network. This is the setup of my network:

1 Network card, Realtek RTL8139 Family PCI Fast Ethernet NIC
1 valid network between the main computer and my friends laptop (ie. Able to connect to network resources, such as \\comp)

My Realtek card is set as such (options may not be available for your card):

No firewall (on windows XP)
Link Speed / Duplex Mode: Auto Mode
Network Address: Not present
Receive buffer size: 64k bytes

And I have the following protocols:

Client for Microsoft Networks (Name Service Provider is "Windows Locator")
File and printer sharing for Microsoft Networks
QoS Packet Scheduler
TCP / IP (IP and DNS Server addresses are obtained automatically, NetBIOS settings is default in "Advanced" settings, LMHosts lookup is on, Connection address registration with DNS is ON)

As long as you have a valid network connection, and can use other chat programs, then NChat should work. NChat has been successfully tested on 5 networks, and on 3 operating systems (XP, 2000 and ME)

If you wish to try NChat on a computer with no network, then do the following:

1) Start NChat and open the options (Main Menu, Options)
2) Click the "Admin" tab
3) Open "<LOCATION OF NCHAT>\Other Stuff\NChat Admin Password Generator\Project1.vbp"
4) Run it and fill in the details
5) Switch to NChat and enter the details
6) Open the "Admin Menu", click Loopback / Broadcast and click YES.
7) NChat should now work on one computer, because it uses the IP address: 127.0.0.1 instead of 255.255.255.255

Q: Why can't I use NChat on computers with 2 active network cards (eg with 2 IP addresses, 169.111.111.111 and 80.1.1.1?
A: The winsock control is VERY basic. It picks the FIRST free Network card found and uses that. Therefore you can't use NChat on computers with 2 cards. To use NChat, use the method outlined in the last question

Q: NChat won't work, and there is no solution in this file
A: Join the forums at www.solidinc.tk. You can get all sorts of help from there.

---------------------------
Section 4: What's Next?
---------------------------

> Web-cam support. How many times have you wanted to show someone something, but they are in another building? Personally, all the time!. So that's why support for web-cams are being added (soon)

> Custom Icon sending. About 90% done. When 100%, will let you use any standard picture (JPG, BMP etc.) as your icon, and everyone will be able to see it, just like on the new MSN Messengers and Trillian.

> More AI for Notch. Notch is stupid. VERY stupid. He needs some kind of intelligence added, maybe contextual linking or something, but all ChatBot AI that I have seen is not suitable for NChat. Maybe i'll program my own... as usual

> S**T LOADS OF CODE STREAMLINING. NChat is very bulky on both hard-drive space, and network bandwidth. Need to slow down on the Broadcast commands, and remove the dirty little "Once-off" hacks and inconsistent coding.

> Thinking of moving the private chat to a more private winsock. Maybe just using UDP, coz it's easier and uses less coding.

> More HTML stuff, like smart tags in MS Word, which will let you query definitions, header stuff etc. 

> Web-Site downloads, such as more profiles, more Notch Scripts, downloadable smiley sets, and HEAPS more!

---------------------------
Section 5: MY CODE!!
---------------------------

A few rumors have surfaced among some people, that this is merely a compilation of codes from www.pscode.com, and that I didn't write much of the code. That is pretty much a lie. While it is true that I did impliment code from other people, I have acknowledged the authors, which is the only condition of using their code. 90% of code on PSCode is code either cleaned up by others, or based on code that others have written. So please, no more nasty comments about my 4 years of hard work