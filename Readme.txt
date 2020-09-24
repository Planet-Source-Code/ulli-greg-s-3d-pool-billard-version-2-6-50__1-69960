GREG'S POOL 3D – PROTOTYPE VERSION.

NOTE: English is not my native language, thus You may find many spelling and gramatic mistakes in this text as well as within the code itself. 

ANOTHER NOTE: I have very little idea about the game of Pool (or Snooker as I think, some tend to call it) and some rules present in this game may be very far from reallity. I hope, that my "interpretation" of snooker will not be offensive to any snooker die-hard... :-)

DISCLAIMER: Though this programme was tested and caused no problemms, I cannot guarantee, that it will run smoothly on Your machine. You use this code on Your own risk.

TO THE POINT - GAME'S DISCRIPTION:
This is a simple pool-like game with Direct3D graphics. It is almost complete and performs generally well (at least on my computer – a 650MHz processor with 192MB of RAM and 16MB on a Riva TNT2 graphics card), though it has some loose-ends (listed later).

The game makes use of basic 3D techniques, like:
-	textured meshes, 
-	vertex and index buffers, 
-	alpha blending, 
-	matrix transformations,
-	billboards 
-	directional lighting

Other built-in features include:
-	2D physics with a collision detection and response mechanism
- 	sprites
-	custom controlls
- 	mobile cameras 
- 	dynamic sound
 	
Known bugs and loose-ends include:
-	Looks well only on 1024x768 resolution
-	Fails, when the Direct3D device is lost
-	Does not have a help system
-	The error handling mechanism works, but is oversimplified
-	The collision detection mechanism tends to fail in certain situations
-	The table could use some details

Game controlls:
The main input device is the mouse. For moving the camera press the left or right mouse button (depending on the type of movement you want) and move the mouse.	There are also few keys that can be used:
Home	- toggles between available cameras
Space	- launches the cue-ball
F2	- starts a new game
F3	- exits
Up Down Left Right PageUp PageDöwn - move the camera

FINAL NOTE: I am not planning on further developping this project. If You like it and have some ideas on how to make it better, feel free to take it and change whatever You like. If you have any questions, comments or remarks feel free to send mail me.

Have fun 
Author: Grzegorz Holdys (Wroclaw, Poland)
E-mail: gregor@kn.pl
