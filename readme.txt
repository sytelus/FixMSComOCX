Fix MSCOM OCX Class IDs In VB Projects 

Call it yet another Microsoft's blunder but you still can't forget when you try to load a VB project and it says something "Can't load MSCOMCTL.OCX" and replaces all your fancy controls with mere picture boxes! Here's small utility that might do the trick. It goes through all your vbp, frm, ctl and replaces GUIDs of crazy OCXes with ones that you have. This application is right now not in state of easy clicks. You will have to find out GUIDs of OCXes installed in your machine and replace it by hand in the source code. However you may freely contact me for assistance. 
 
This Application and it's source code is copyrighted to,
(C) Shital Shah, 1998-2001.
www.ShitalShah.com
email: shital@ShitalShah.com