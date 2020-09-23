Visual Dialog++ Readme File
Version 1.0.0 BETA 2
Copyright © 2003, Shukri Zahari



Introduction

Nowday, a visual designer for a particular programming language is very important. Without a visual designer, you have to code every little thing about your application (that take a long journey). I've seen one visual designer a 80% like VB one ( I don't remember where I found it...) but the code use a lot of subclass. Subclass is not a very good way to code your program as it can crash your VB or the ENTIRE system!! So, in this past few week, I found some articles & samples on book I've bought & decide to make my own visual designer & here you go Visual Dialog++, the safe way to move control on your form. This code I think will work great with JEREMY BOYD's Visual Sight++. You can found his code at:

http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43397&lngWId=1

or just go to his website at: HTTP://DSS1.CJB.NET




Disclaimer of Warranty

I, hereby as the developer of this code didn't take responsibility of anything happen to you, your PC, or software. You must use this codes at your own risk.



Update Info

- No more flickering grid when you move the form by directly paint the grid picture to the form using PaintPicture method
- When using WinXP, there will be additional blue border surrounding the CommandButton. The border removed when you release your mouse by applying the SendMessage API (see FORM_LOAD procedure for details...)
- Visual Dialog++ now can add control dynamically using control array at runtime.
- Experimenting with resizing the arrayed control at runtime (not finished yet... :-P
- Sleek & clean new GUI compared to version 1 BETA 1 (Thanks to my friend, Lesnar for helping me out...)
- Right-Click on control will set the selected control's ZOrder to 0



Why Visual Dialog++?

i  - Advantages
     - No more subclassing method that can crash not only VB but your entire system
     - No additional module or class module. All have been included in one form
     - 100% pure VB+API

ii - Easy to use
     - You can apply technics used easily on your own application
     - Codes are royalty-free



Unsolved Problems

- Resize the arrayed control at runtime: It happen to be an error when I try to resize it. See the ResizeCTL function
- How could I move Frame? Although Frame has the Window Handle (hWnd), but I can't move it. Why????????
- When using VB default controls (eg. label, image etc...), I could only move 8 out of 18 controls. Any idea to add the other controls???
- Why the Windows XP manifest file embedded in a RES couldn't work properly? Sometime work, sometime not...



Limitations:

- You can use this codes freely everywhere
- You must include me in the credit



Feedback:

- You could email me at: FredrickDallasDurst@yahoo.com
                         programmingdaily@yahoo.com
- Leave comments on PSC.com
- Vote me & give at least 1 globe (it should be enough...)
- If you has a WinXP machine, try compile the code and run the EXE. Tell me if they can run on your PC (It won't run on mine)




PS: The codes maybe messed up but I don't have time to reorganize it coz I had an MCSE exam this month

PS: Who interested in making a form designer like VB, you can share your knowledges about API with me. I'm a muthafucking VB novice :-#

