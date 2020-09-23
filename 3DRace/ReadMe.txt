Hello All.

Time ago, my nephew asked me if was possible to make some modifications to a game he loved 
to play called LUDORACE: built new maps (it comes without a map editor,sounds and music),
change the settings of the player and of the opponent racers, add music and sounds and so on.

Unfortunatly, the only way to do that was to re-built the game (I think).

Therefore I began to work to this project in my spare time;
but now, my nephew is falling in love for other new games and he has discarded the above
game (and the mine too).

So I have decided to upload my application to you.

The game is playable but not complete (it lacks a good intro, a menu to open maps, options,
relative position, overlap, ...), it has some tricks and quirks (like as final ranks, negative 
velocity, ...) and there are some little inconvenients, at today:
.player don't has acceleration and the velocity result also negative;
.the raypick collision detection is good if you go forward but it's not perfect in backward
 (probably I will change this method with a classical 2D collision detection);
.the velocity of player and of opponent racers is not set with the system clock (so increase it
 on slow pc or decrease it on powerful computers);
.the camera and opponent drivers not have collision detection with the walls;
.the result of collision between the racers is not very accurate (it tests only the velocity) 
.the drivers have an awfull lateral vibration when running;
.probably some others I don't remember at this moment.

Try to solve these problems and to complete the game if you like. 

This program was written in VB6 and it's using DirectX7 D3D Retain Mode but it's possible to use
DX7/DX8 Immediate Mode with few modifications of the code, if you prefer.

Let me know if you like it (via e-mail or voting);
if so and if you will be a lot, I'll try to finish it and I'll improve my map editor to upload 
it (it's too slow and full of tricks and quirks to upload it now).

Everybody can modify the code or employ parts of it within own projects, but please let me
know about your projects and give me a little credit.
Only one restriction: NOBODY MUST USE ANY PART OF THIS PROGRAM IN COMMERCIAL PURPOSES.

I wish to thank some programmers for their code examples, from which I have learned a lot 
about D3D (without them this program could not be realized):
  
K. Sudhakar for his DX7RM engine,
N. Turner for his AI racers sample,
W. Kienreich for his raypick method rule

and thanks to GraphitX and Ludorama for their original game.


Enjoy this amazing app!

I would appreciate it if you give me feedback.

Every comments, suggestions, ideas and e-mails are always welcomed to:
fabiocalvi@yahoo.com


How to play
-----------

It's very easy:

Arrow up key => go forward
Arrow down key => go backward
Arrow left key => turn left
Arrow right key => turn right
Escape key => quit the game
V key => backward view
F1 to F12 => fixed camera views along the route

When you pick up a speed bonus your velocity (max 10) increases of 1 unit until holding on 
the arrow up key pressed, but becomes normal if you stop to press the arrow key up or collide 
against the walls or fence.

Bye, Fabio.

P.S.:
on slow computers the initial loading of meshes could be take a long time, so be patience;
hardware acceleration is not necessary but suggested.
