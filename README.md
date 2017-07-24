#   MutiEditor for Wario Land 4 is being made with VB
## PS: patch applications will be made with C++
## 
##     what it can do now:
##      1. The App read the Levels and its rooms data directly from the ROM file (*.gba for wario land 4) and you can use the app to edit one layer of a room by editing textmap value or by using visual MAP editor to make change in visual room rendered out. And also you can change 3 layer at a same time and save them all in visual MAP editor.
##      2. the visual map editor haven't been finished yet and it only can render a whole room without alpha blending. (I mean sometimes you will find some of the layer 0 just cover the whole area in one color and there is some holes on it, if you don't know how it works, just uncheck the checkbox for layer 0 and refresh the MAP. )
##      3. there are some other things the App can do, but are only for an alpha test, without WL4 hacking information, it will be hard to use them properly when you try to edit enimies and timers etc. I will try to make wizards if any part of hacking is completed.
##      4. Camera control modification is available in Visual MAP Editor visually, there are two types available and the textbox just show the exact control characters in the ROM file.
##      5. You can Edit Sprites Tiles by input a Index. but only bitmap-direct-change available. The sprites editor sometimes will not render the sprites in right colors, you need to use the slider to correct it.
##      6. Sprites stuffs and their positions can only change in textmode, it's not wise to change them if you don't know how the characters function. By the way, there will be a lot of work if I make a wizard for this, so it will be a long way to get there. 
##
##      Although VB6 App can run on various Windows platform, but due to I coding on Windows 10 Creater with a HD wide screen, the App UI always not suitable for average screen. I am trying to fix the problem now.
##
##      I am going to make a release version when I finished the inportant parts of the App and make it stable enougth on running. And this is a view for the App when editting with the Visual MAP Editor:
![Image text](https://github.com/shinespeciall/WarioLand4MultiEditor/blob/master/App_Screenshot.png)
##      If you want to rebuild the level but not for a test (now you can find suitable Tile textcode for yourself in the room or other rooms use the same tileset to make suitable changes Or try to use the Visual editor), you need to expand your source gba file, which is available as a patch program here now, rewrite data usually need more room.
##      This is an example for what it can do:
![Image text](https://github.com/shinespeciall/WarioLand4MultiEditor/blob/master/screenshot.png)
