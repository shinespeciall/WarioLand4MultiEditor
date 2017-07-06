#   MutiEditor for Wario Land 4 is being made with VB
## PS: patch applications will be made with C++
## 
##     what it can do now:
##      1. The App read the Levels and its rooms data directly from the ROM file (*.gba for wario land 4) and you can use the app to edit one layer of a room by editing textmap value or by using visual editor to make change in visual room rendered out. And also you can change 3 layer at a same time and save them all in newly build version.
##      2. the visual map editor haven't been finished yet and it only can render a whole room without alpha blending. (I mean sometimes you will find some of the layer 0 just cover the whole area in one color and there is some holes on it, if you don't know how it works, just uncheck the checkbox for layer 0 and refresh the MAP. )
##      3. there are some other things the App can do, but are only for an alpha test, without WL4 hacking information, it will be hard to use them properly when you try to edit enimies and timers etc. I will try to make wizards if any part of hacking is completed.
##      4. Camera control modification is available in Visual MAP Editor visually, there are two types available and the textbox just show the exact control characters in the ROM file.
##      5. You can Edit Sprites Tiles by input a Index. but only bitmap-direct-change available. The sprites editor sometimes will not render the sprites in right colors, you need to use the slider to correct it.
##      6. Sprites stuffs and their positions can only change in textmode, it's not wise to change them if you don't know how the characters function. By the way, Its will be a lot of work if I make a wizard for this, so it will be a long way to get there. 
##
##      I am going to make a release version when I finished some inportant parts of the App and make it stable enougth on running.
##      If you want to run the App hurriedly, you will need COMDLG32.OCX (x86) and comctl32.ocx (x86) to run the App properly.(download them and put them in the same path of the App, the files are in old version so I suppose you to download them althougth they are contained in any of your Microsoft .NET framework.)
##    I recently change the language from Chinese to English for publishing, all the important parts has been translated but there are something excluded and some bad expressions haven't been fixed. so until now, not all the messages are showed in English.
##      If you want to rebuild the level but not for a test (now you can only find suitable Tile text for yourself in the room or other rooms use the same tileset to make suitable change Or try to use the incompleted Visual editor), you need to expand your source gba file, which is available as a patch program here now, rewrite data usually need more room.
##      This is an example for what it can do:
![Image text](https://github.com/shinespeciall/WarioLand4MultiEditor/blob/master/screenshot.png)
## 
## for those who want to remake the game
### I just want to remake the game in level in origin ROM, it is just an interest. If you want to ask what it will be like, I will mention another popular game Super Mario Maker as an example. with enougth things you can operate, you are able to remake the game.
### Its just a Wario Land 4 fan's homemake repo for all my project, I won't change the charactor or the game mode. I suppose you to make your own version for your own taste. If you would like to make your own game version by this App and want to improve the present working condition, you can go to CONTRIBUTING.md or help me something in folder what_I_am_doing. This App update and build frequently and you are supposed to download newly build version weekly.
