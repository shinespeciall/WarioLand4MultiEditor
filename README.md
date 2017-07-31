#   MultiEditor for Wario Land 4 is being made with VB
## PS: patch applications will be made with C++
## 
##     what it can do now:
##      1. The App read the Levels and its rooms data directly from the ROM file (*.gba for wario land 4) and you can use the app to edit one layer of a room by editing textmap value or by using visual MAP editor to make change in visual room rendered out. And also you can change 3 layers at the same time and save them all in visual MAP editor.
##      2. The visual map editor have almost been finished and it recently can render a whole room with alpha blending. (if you don't know how it works, just uncheck the checkbox named Alpha and refresh the MAP. ) And the most important thing is that you can make change directly on the MAP with more than one Layer rendering on the board, you just need to know the change will be made on the front layer which is being rendering now.
##      3. there are some other things the App can do, but are only for an alpha test, without WL4 hacking information, it will be hard to use them properly when you try to edit enimies and timers etc. I will try to make wizards if any part of hacking is completed.
##      4. Camera control modification is available in Visual MAP Editor visually, there are two types available and the textbox just show the exact control characters in the ROM file.
##      5. You can Edit Sprites Tiles by input a Index. but only bitmap-direct-change available. The sprites editor sometimes will not render the sprites in right colors, you need to use the combobox to correct it.
##      6. Sprites stuffs and their positions can only change in textmode, it's not wise to change them if you don't know how the characters function. By the way, there will be a lot of work if I make a wizard for this, so it will be a long way to get there. 
##      7. the txt files in directory "~\MOD\ " are models for Tilesets which will be loaded into running App, it can be created or changed automatically by the App if you try to edit a room using a particular Tileset or make models for it visually in the Visual MAP Editor, sometimes files will be changed or uploaded if I make some test in Visual MAP Editor, I will be happy to see you contribute to this folder.
##
##      Although VB6 App can run on various Windows platform, but due to I coding on a HD wide screen, the App UI always not suitable for average screen. I have tried to solve the problem and if bugs still exist please make an issue here.
##
##      I am going to make a release version when I finished some important parts of the App, if you want to run the alpha version, just download the whole repo and run the WL4 MutiEditor.exe as administrator. You can report bugs by creating an issue to the repo and this is a screenshot for the App when editting with the Visual MAP Editor:
![Image text](https://github.com/shinespeciall/WarioLand4MultiEditor/blob/master/App_Screenshot.png)
##      This is an example for what it can do:
![Image text](https://github.com/shinespeciall/WarioLand4MultiEditor/blob/master/screenshot.png)
##      PS: If you want to rebuild the level but not for a test (now you can find suitable Tile textcode for yourself in the room or other rooms use the same tileset to make suitable changes Or try to use the Visual editor), you need to expand your source gba file, which is available as a patch program here now, rewrite data usually need more room.
