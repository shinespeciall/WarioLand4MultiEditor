VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "New Room Wizard"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   15585
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   11640
      TabIndex        =   9
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Layers' Attributes"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   11175
      Begin VB.CheckBox Check1 
         Caption         =   "using Layer(0) to render Smoke"
         Height          =   375
         Left            =   7680
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Text            =   "< Choose a way for alpha blending >"
         Top             =   960
         Width           =   6615
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Text            =   "< Choose a priority order >"
         ToolTipText     =   "Choose Tileset and BG MAP first"
         Top             =   360
         Width           =   10575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "BG MAP"
      Height          =   7575
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   7935
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   6375
         Left            =   240
         ScaleHeight     =   6315
         ScaleWidth      =   7395
         TabIndex        =   5
         Top             =   960
         Width           =   7455
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Text            =   "<Choose one existent BG MAP>"
         Top             =   360
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tileset"
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Text            =   "00  Debug room"
         Top             =   360
         Width           =   6375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   6375
         Left            =   240
         ScaleHeight     =   6315
         ScaleWidth      =   6315
         TabIndex        =   1
         Top             =   960
         Width           =   6375
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Form11.Enabled = False
Form11.Picture1.Cls
Form11.Picture2.Cls
Form11.Combo2.Clear
Form11.Combo2.Text = "<Choose one existent BG MAP>"
Form11.Combo2.ListIndex = -1

Form11.Check1.Value = 0
If Form11.Combo1.ListIndex = Val("&H21") Or Form11.Combo1.ListIndex = Val("&H22") Then
Form11.Check1.Visible = True
Else
Form11.Check1.Visible = False
End If

Dim Tilesets As String
Dim StrTemp As String, str1 As String, str2 As String, str3 As String, str4 As String
Dim TileOffset As String, TileLength2 As Long
Dim TextMAPDataOffset As String, paletteOffset As String
Form9.Text1.Text = Form9.Text1.Text & "impoting Tileset......" & vbCrLf
Tilesets = Mid$(Form11.Combo1.Text, 1, 2)
StrTemp = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 9 * 4 + Val("&H" & "3F2298")), Hex(Val("&H" & Tilesets) * 9 * 4 + 35 + Val("&H" & "3F2298")))
TileOffset = Mid$(StrTemp, 5, 2) & Mid$(StrTemp, 3, 2) & Mid$(StrTemp, 1, 2)
TileLength2 = Val("&H" & Mid$(StrTemp, 11, 2) & Mid$(StrTemp, 9, 2))
Form9.Text1.Text = Form9.Text1.Text & "The amount of Tiles is " & str(TileLength2 / 32) & vbCrLf
TextMAPDataOffset = Mid$(StrTemp, 45, 2) & Mid$(StrTemp, 43, 2) & Mid$(StrTemp, 41, 2)
paletteOffset = Mid$(StrTemp, 21, 2) & Mid$(StrTemp, 19, 2) & Mid$(StrTemp, 17, 2)
ReDim Palette256(16, 16)
ReDim Tile88(2048)
Form9.Text1.Text = Form9.Text1.Text & "impoting 8 * 8 Tiles Data......" & vbCrLf
Dim TextMapData As String
StrTemp = ReadFileHexWithByteInterchange(gbafilepath, TileOffset, Hex(Val("&H" & TileOffset) + TileLength2))
'-------------------------------------------------------------权宜之策
str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F8098")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F8098")))
'str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F91D8")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F91D8")))             '???
For i = 0 To 15
str2 = Mid$(str1, 4 * i + 1, 2)
Mid$(str1, 4 * i + 1, 2) = Mid$(str1, 4 * i + 3, 2)
Mid$(str1, 4 * i + 3, 2) = str2
Next i
For i = 0 To 15
str2 = Mid$(str1, 4 * i + 1, 4)
str3 = ReadFileHex(gbafilepath, Hex(Val("&H" & str2) * 8 + Val("&H" & "3F7828")), Hex(Val("&H" & str2) * 8 + Val("&H" & "3F782F")))
str2 = Mid$(str3, 9, 8)
str4 = Mid$(str3, 1, 2)
str3 = Mid$(str3, 5, 2)
If str4 = "03" Or str4 = "06" Then
str3 = Hex(Val("&H" & str3) - 1)
Else
str3 = "00"
End If
str2 = Mid$(str2, 7, 2) & Mid$(str2, 5, 2) & Mid$(str2, 3, 2) & Mid$(str2, 1, 2)
str2 = Hex(Val("&H" & str2) - Val("&H8000000") + Val("&H" & str3) * 128)
str2 = ReadFileHexWithByteInterchange(gbafilepath, str2, Hex(Val("&H" & str2) + 127))
DoEvents
For j = 0 To 3
Tile88(i * 4 + j) = Mid$(str2, j * 64 + 1, 64)
Next j
Next i
Tile88(64) = Replace(Space(64), Chr(32), "0")
For i = 0 To (TileLength2 / 32) - 1
Tile88(i + 65) = Mid$(StrTemp, 64 * i + 1, 64)
DoEvents
Next i
For i = (TileLength2 / 32 + 65) To 2047
Tile88(i) = Replace(Space(64), Chr(32), "0")
DoEvents
Next i
'-------------------------------------------------------------End  权宜之策

Form9.Text1.Text = Form9.Text1.Text & "impoting and making 16 * 16 Tiles Data......" & vbCrLf
TextMapData = ReadFileHex(gbafilepath, TextMAPDataOffset, Hex(Val("&H" & TextMAPDataOffset) + 8192))
For i = 0 To 4095
StrTemp = Mid$(TextMapData, 4 * i + 1, 2)
Mid$(TextMapData, 4 * i + 1, 2) = Mid$(TextMapData, 4 * i + 3, 2)
Mid$(TextMapData, 4 * i + 3, 2) = StrTemp
DoEvents
Next i

ReDim Tile16(1024)
Dim r0 As Long, r1 As Long, r2 As Long
For i = 0 To 1023
r0 = i * 4
r2 = r0 Or 1
Tile16(i) = Mid$(TextMapData, r0 * 2 * 2 + 1, 4)
r1 = (r2 + 1) * 2 ^ 16
Tile16(i) = Tile16(i) & Mid$(TextMapData, r2 * 2 * 2 + 1, 4)
r0 = 128 * 2 ^ 9
r2 = r1 + r0
r1 = RSH(r1, 15)
Tile16(i) = Tile16(i) & Mid$(TextMapData, r1 * 2 + 1, 4)
r2 = RSH(r2, 15)
Tile16(i) = Tile16(i) & Mid$(TextMapData, r2 * 2 + 1, 4)
DoEvents
Next i

Form9.Text1.Text = Form9.Text1.Text & "impoting palette 256 ......" & vbCrLf
StrTemp = ReadFileHex(gbafilepath, paletteOffset, Hex(Val("&H" & paletteOffset) + 256 * 2 - 1))
For j = 0 To 15
For i = 0 To 15
Palette256(i, j) = RGB555ToRGB888(Mid$(StrTemp, 64 * j + 4 * i + 1, 4))
DoEvents
Next i
Next j
Dim a As Boolean
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i, j, Hex(i + 8 * j), Form11.Picture1, False, 24)
DoEvents
Next i
Next j
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i + 8, j, Hex(128 + i + 8 * j), Form11.Picture1, False, 24)
DoEvents
Next i
Next j

Dim BGMAPpath As String
BGMAPpath = App.Path & "\MOD\" & Tilesets & " BGMAPDATA.txt"
If Dir(BGMAPpath) = "" Then
    Open BGMAPpath For Append As #4
    Print #4, "Universal Blank BG"
    Print #4, "000858DA7C";
    Close #4
End If

Open BGMAPpath For Input As #4
ReDim BGMAPHeader(1, 10)
i = 0
Do While Not EOF(4)
    Line Input #4, BGMAPHeader(0, i)
    Form11.Combo2.AddItem BGMAPHeader(0, i)
    Line Input #4, BGMAPHeader(1, i)
i = i + 1
Loop
Close #4
'Erase Palette256(16, 16)    'used again in BG Mapping
Erase Tile88()
Erase Tile16()
Form11.Enabled = True
Form11.Combo2.Enabled = True
End Sub

Private Sub Combo2_Click()
Form11.Enabled = False
Dim str As String, StrTemp As String
Dim i As Long, j As Long
Form11.Picture2.Cls
If Mid$(BGMAPHeader(1, Form11.Combo2.ListIndex), 1, 2) = "00" Then Exit Sub

str = Mid$(BGMAPHeader(1, Form11.Combo2.ListIndex), 3, 8)
str = Hex(CLng("&H" & str) - CLng("&H8000000"))
str = DecompressRLE(str)
If str = "" Then
MsgBox "Something Wrong when Decompressing !"
layerHeight = 0
layerWidth = 0
Erase TextMap()
Exit Sub
End If
If Mid$(BGMAPHeader(1, Form11.Combo2.ListIndex), 1, 2) = "10" Then
MsgBox "Something Wrong with BGMAPDATA File !"
layerHeight = 0
layerWidth = 0
Erase TextMap()
Exit Sub
End If

ReDim Tile88(1023)
For i = 0 To 1023
Tile88(i) = Replace(Space(64), Chr(32), "0")
Next i

Dim BGTileOffset As String, BGTileLength As Long, Tilesets As String

Tilesets = Mid$(Form11.Combo1.Text, 1, 2)
StrTemp = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 9 * 4 + CLng("&H" & "3F2298")), Hex(CLng("&H" & Tilesets) * 9 * 4 + 35 + CLng("&H" & "3F2298")))   'Get entirety of them
BGTileOffset = Mid$(StrTemp, 29, 2) & Mid$(StrTemp, 27, 2) & Mid$(StrTemp, 25, 2)
BGTileLength = CLng("&H" & Mid$(StrTemp, 35, 2) & Mid$(StrTemp, 33, 2))
StrTemp = ReadFileHexWithByteInterchange(gbafilepath, BGTileOffset, Hex(CLng("&H" & BGTileOffset) + BGTileLength))

j = 0
For i = (1023 - (BGTileLength / 32)) To 1022
Tile88(i) = Mid$(StrTemp, 64 * j + 1, 64)
j = j + 1
Next i

For j = 0 To Min(layerHeight - 1, Form11.Picture2.Height \ 2 - 1)
For i = 0 To Min(layerWidth - 1, Form11.Picture2.Width \ 2 - 1)
DrawTile8 i, j, TextMap(i, j), Form11.Picture2, , 14
Next i
Next j
layerHeight = 0
layerWidth = 0
Erase TextMap()
Erase Tile88()
Form11.Enabled = True
Form11.Frame3.Enabled = True
End Sub

Private Sub Combo3_Click()
Form11.Combo4.Enabled = True
End Sub

Private Sub Combo4_Click()
Form11.Command1.Enabled = True
End Sub

Private Sub Command1_Click()
If LevelStartStreamOffset = "" Then Exit Sub
If SaveDataOffset(95) <> "" Then
    MsgBox "buffer memory used up, save all and retry !"
    Exit Sub
End If
Dim i As Integer, j As Integer    ', str1 As String
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
Next i
SaveDataOffset(i) = Hex(Val("&H" & LevelStartStreamOffset) + 1)    '修改Room数量标志位，最大值为 10 Hex
SaveDatabuffer(i) = Right("00" & Hex(Val("&H" & RoomNumber) + 1), 2)
RoomNumber = Right("00" & Hex(Val("&H" & RoomNumber) + 1), 2)
SaveDataOffset(i + 1) = LevelAllRoomPointerandDataBaseOffset         '每个Room的layer指针和元素指针及Flag数据串保存基址

SaveDatabuffer(i + 1) = LevelAllRoomPointerandDataallHex & Right("00" & Hex(Form11.Combo1.ListIndex), 2) & "101010 20000000 63223F08 63223F08 63223F08 7CDA5808"     'Normal

End Sub

Private Sub Form_Load()
Form11.Width = 15825
Form11.Height = 10180
Form11.Left = Form4.Width
Form11.Icon = LoadResPicture(101, vbResIcon)
Form11.Top = 0
Form11.Combo1.FontSize = 15
Form11.Combo2.FontSize = 15
Form11.Combo3.FontSize = 15
Form11.Combo4.FontSize = 15
Form11.Check1.Value = 0
Form11.Check1.Visible = False
Form11.Command1.FontSize = 17
Form11.Combo1.AddItem "00  Debug room"
Form11.Combo1.AddItem "01  Palm Tree Paradise"
Form11.Combo1.AddItem "02  Caves"
Form11.Combo1.AddItem "03  The Big Board"
Form11.Combo1.AddItem "04  The Big Board"
Form11.Combo1.AddItem "05  The Big Board?"
Form11.Combo1.AddItem "06  Wildflower Fields"
Form11.Combo1.AddItem "07  Toy Block Tower"
Form11.Combo1.AddItem "08  Factory"
Form11.Combo1.AddItem "09  Wildflower Underground"
Form11.Combo1.AddItem "0A  Wildflower WaterPlace"
Form11.Combo1.AddItem "0B  Underwater"
Form11.Combo1.AddItem "0C  Toy Block Tower"
Form11.Combo1.AddItem "0D  Toy Block Tower"
Form11.Combo1.AddItem "0E  Toy Block Tower"
Form11.Combo1.AddItem "0F  Doodle"
Form11.Combo1.AddItem "10  Dominoes"
Form11.Combo1.AddItem "11  Hall of Hieroglyphs"
Form11.Combo1.AddItem "12  Haunte House (plus debug tiles)"
Form11.Combo1.AddItem "13  Crescent Moon Village outside"
Form11.Combo1.AddItem "14  Drain"
Form11.Combo1.AddItem "15  Arabian outside"
Form11.Combo1.AddItem "16  Arabian inside"
Form11.Combo1.AddItem "17  Arabian"
Form11.Combo1.AddItem "18  Arabian"
Form11.Combo1.AddItem "19  Arabian"
Form11.Combo1.AddItem "1A  Dominoes (blue)"
Form11.Combo1.AddItem "1B  Dominoes (purple)"
Form11.Combo1.AddItem "1C  Dominoes (teal)"
Form11.Combo1.AddItem "1D  Factory"
Form11.Combo1.AddItem "1E  Factory"
Form11.Combo1.AddItem "1F  Jungle"
Form11.Combo1.AddItem "20  Factory"
Form11.Combo1.AddItem "21  Toxic Landfill"
Form11.Combo1.AddItem "22  Toxic Landfill"
Form11.Combo1.AddItem "23  Pinball"
Form11.Combo1.AddItem "24  Pinball"
Form11.Combo1.AddItem "25  Pinball (with Gorilla)"
Form11.Combo1.AddItem "26  Jungle"
Form11.Combo1.AddItem "27  40 Below Fridge"
Form11.Combo1.AddItem "28  Jungle"
Form11.Combo1.AddItem "29  Jungle caves"
Form11.Combo1.AddItem "2A  Hotel"
Form11.Combo1.AddItem "2B  Hotel"
Form11.Combo1.AddItem "2C  Hotel"
Form11.Combo1.AddItem "2D  Hotel"
Form11.Combo1.AddItem "2E  Hotel"
Form11.Combo1.AddItem "2F  Hotel (outside)"
Form11.Combo1.AddItem "30  Unused in-game (Haunted House)"
Form11.Combo1.AddItem "31  Unused in-game (Haunted House)"
Form11.Combo1.AddItem "32  Unused in-game (Cardboard)"
Form11.Combo1.AddItem "33  Cardboard"
Form11.Combo1.AddItem "34  Caves"
Form11.Combo1.AddItem "35  Jungle"
Form11.Combo1.AddItem "36  Caves"
Form11.Combo1.AddItem "37  Lava level"
Form11.Combo1.AddItem "38  Caves"
Form11.Combo1.AddItem "39  Golden Passage"
Form11.Combo1.AddItem "3A  Hotel"
Form11.Combo1.AddItem "3B  Hotel"
Form11.Combo1.AddItem "3C  Hotel"
Form11.Combo1.AddItem "3D  Hotel"
Form11.Combo1.AddItem "3E  40 Below Fridge"
Form11.Combo1.AddItem "3F  Factory"
Form11.Combo1.AddItem "40  Factory"
Form11.Combo1.AddItem "41  Arabian"
Form11.Combo1.AddItem "42  Boss room"
Form11.Combo1.AddItem "43  Boss corridor"
Form11.Combo1.AddItem "44  Boss room"
Form11.Combo1.AddItem "45  Frozen lava level"
Form11.Combo1.AddItem "46  Lava level"
Form11.Combo1.AddItem "47  Hall of Hieroglyphs"
Form11.Combo1.AddItem "48  Boss room"
Form11.Combo1.AddItem "49  Boss room"
Form11.Combo1.AddItem "4A  Boss corridor"
Form11.Combo1.AddItem "4B  Boss corridor"
Form11.Combo1.AddItem "4C  Boss corridor"
Form11.Combo1.AddItem "4D  Boss corridor"
Form11.Combo1.AddItem "4E  Boss corridor"
Form11.Combo1.AddItem "4F  Boss room (Diva)"
Form11.Combo1.AddItem "50  Hall of Hieroglyphs"
Form11.Combo1.AddItem "51  Jungle"
Form11.Combo1.AddItem "52  Wildflower"
Form11.Combo1.AddItem "53  Crescent Moon Village"
Form11.Combo1.AddItem "54  Crescent Moon Village"
Form11.Combo1.AddItem "55  Crescent Moon Village"
Form11.Combo1.AddItem "56  Toy Block Tower"
Form11.Combo1.AddItem "57  Pinball"
Form11.Combo1.AddItem "58  Bonus room"
Form11.Combo1.AddItem "59  Bonus room"
Form11.Combo1.AddItem "5A  Final level"
Form11.Combo1.AddItem "5B  The Big Board end"
Form11.Combo3.AddItem "layer(0) Priority = 0: layer(1) Priority = 1: layer(2) Priority = 2"
Form11.Combo3.AddItem "layer(0) Priority = 1: layer(1) Priority = 0: layer(2) Priority = 2"
Form11.Combo3.AddItem "layer(0) Priority = 1: layer(1) Priority = 0: layer(2) Priority = 2"
Form11.Combo3.AddItem "layer(0) Priority = 2: layer(1) Priority = 0: layer(2) Priority = 1"
Form11.Combo4.AddItem "No Alpha Blending"
Form11.Combo4.AddItem "EVA = 7: EVB = 16"
Form11.Combo4.AddItem "EVA = 10: EVB = 16"
Form11.Combo4.AddItem "EVA = 13: EVB = 16"
Form11.Combo4.AddItem "EVA = 16: EVB = 16"
Form11.Combo4.AddItem "EVA = 16: EVB = 0"
Form11.Combo4.AddItem "EVA = 13: EVB = 3"
Form11.Combo4.AddItem "EVA = 10: EVB = 6"
Form11.Combo4.AddItem "EVA = 7: EVB = 9"
Form11.Combo4.AddItem "EVA = 5: EVB = 11"
Form11.Combo4.AddItem "EVA = 3: EVB = 13"
Form11.Combo4.AddItem "EVA = 0: EVB = 16"
Form11.Combo4.AddItem "EVA = 0: EVB = 16"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase BGMAPHeader()
Erase Palette256()
End Sub
