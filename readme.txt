 SocketWrench File Transfer Example (VB6 SP4)
 Written by Arthur Nisnevich [optik_burner]
 Version 1.0 - 13 Dec. 2000

 DESCRIPTION

 This is an example of how to utilze the Catalyst SocketWrench ActiveX control
 to send and receive remote files between client and server. While an FTP example
 already exists, I have not yet found any good examples of file transfer between
 client server, thus I wrote this.

 IDE INFORMATION

 This application was written in Visual Basic 6.0 SP4 on a Windows 2000 machine.
 It was tested on most Windows 32-bit operating systems. It was not tested on any
 16-bit operating systems. It is recommended that you run this program only on a
 true IBM compatible (not "virtual" Windows). 

 The Catalyst SocketWrench ActiveX control must be installed on your system.
 Visit the Catalyst website below for more information on this.

 <http://www.catalyst.com>

 If you are experiencing any problems getting this program to load, compile, or
 execute, please contact me at <maxmouse@iquest.net> with specific information
 about your problem.

 SOURCE/PROGRAM DOCUMENTATION

   I.	Introduction
   II.  How To Get This Working
   III. Bugs/Issues/Notes
   IV.   Contact Me

 I. INTRODUCTION
 
 This was written for the sole purpose of informing other programmers how to
 transfer files between client and server using Catalyst SocketWrench control,
 a powerful and unique replacement to Winsock ActiveX. 

 Currently, this features basic file transfer (works with small and large files,
 with no limit), buffering, transfer cancelation, progress monitor (how to monitor
 the progress of the file transfer), and various basic SW techniques.

 Feel free to distribute this program and source code as you wish. I am not
 requiring any credit, copyright, or notice placed on your programs, but feel free
 to provide my email and name if you like. :) If you are still experiencing trouble,
 I highly encourage you to email me or reply to my post on any VB developer site.

 Enjoy the code,
 Arthur Nisnevich 
 aka. optik_burner

 II. HOW TO GET THIS WORKING
 
 First off, you will need a copy of Catalyst SocketWrench. At time of this release
 it is version 3.5. You can download the control from:

 <http://www.catalyst.com>

 To get this fully working you have to have 2 instances of the program. Basically,
 just run the program twice. On one of the instances find the "Server" frame on the
 right, and click "Begin Listening...". You can also set a custom port other than
 the default provided, if you like. 

 Then, on the second instance of the program, find the "Client" frame on the
 left, and click "Connect..." Make sure the IP is set to either "localhost", your
 machine name (prefixed by two slashes: \\), or "127.0.0.1". Also, make sure the
 "Port" number matches the one you typed in on the first instance.

 On both of the instances in their specific "Status" frames, you should see the text:

 "Connected. Ready."

 If you do not see the text and you see "Listening..." or "Connecting..." then you have
 forgotten a step. Go back to the beginning and try again.

 You are now connected to yourself, and are ready to send files. In the first instance
 (which we will now reffer to as the Server), select a File Receival path. The default
 one ("(App Path)") symbolizes the path in which the executable is found (in VB terms
 that's App.Path). A second default value is provided, "C:\", which is the standard
 C: drive. Click on "Browse" to select your own path.

 From the first instance (which we will know reffer to as the Client), browse for or
 type in a filename. Click "Send File" when you're ready. Walla! Your file is instantly
 being sent!

 Now I encourage you to take a look at the source code, which is heavily commented.

 III. BUGS, ISSUES, NOTES

 (1) Setting a Buffer Speed
 - If you are experiencing a slow send, try changing the Buffer Speed, which is at
 the "Initial Options" frame at the bottom of the dialog. If you are still experiencing
 slow transfer when using the "T1/LAN >" preset buffer, try setting a custom buffer.
 Remember, the buffer is written in bytes.

 (2) Buffer Speeds "xDSL" to "T1" are the same!
 - Yes, it's true. They are the same. That's because anything above ISDN will generate
 an Overflow error. I am still not exactly sure what is causing this.

 (3) After Canceling a File Transfer on the Server, the Client is Still Going
 - Yep, I'm aware of that, too. I have added extra procudures to Disconnect() events of
 the client, however I still cannot figure out what's causing it to send. Perhaps, the
 Exit Do commands aren't being called. I'll keep at it...
 - This problem does not occur when canceling the file on the Client.

 (4) Error Handling
 - Because this code is for Intermediate programmers, I have not added heavy error
 handling procedures. For example, typing in an invalid path will generate an error when
 trying to write to the file. 
 - I have, however, added important error handling, such as return value testing, file
 existance, port/ip checking, buffer checking, and several others.
 - If you run the program as it should be run, you should not receive any errors, however
 playing with values and variables could generate a possible error. 
 
 IV. CONTACT ME

 If for some reason you need to contact me, use the following email addresses:

 (1) maxmouse@iquest.net -- Primary
 (2) optik@g33k.net	 -- Secondary

 Replying to a post anywhere I submit will automatically send me an email notification, 
 so that will work, too. PLEASE, I REPEAT, PLEASE, NO BASIC QUESTIONS - REMEMBER THIS
 CODE IS INTENDED FOR INTERMEDIATE PROGRAMMERS! I will try to respond to all of your
 queries, but I will simply ignore any basic, stupid, or ridiculous questions.

 [ This code is being put into work in Eclipse(R), a new P2P communications tool. ]
 [ http://eclipsecentral.hypermart.net ] - [ Contact me for Positions & Questions ]

 ===
 EOF