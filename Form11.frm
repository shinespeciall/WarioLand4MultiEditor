VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "New Room Wizard"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   15585
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   495
      Left            =   12000
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tileset (Only show partly just for choose)"
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Text            =   "00  Debug room"
         Top             =   360
         Width           =   6255
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
Form11.Picture1.Cls
If gbafilepath = "" Then
MsgBox "No GBA file Loaded", vbInformation, "Info"
Form11.Visible = False
Exit Sub
End If
Dim Tilesets As String
Dim StrTemp As String, str1 As String, str2 As String, str3 As String, str4 As String
Dim TileOffset As String, TileLength2 As Long
Dim TextMAPDataOffset As String, paletteOffset As String
Form9.Text1.Text = Form9.Text1.Text & "impoting Tileset......" & vbCrLf
Tilesets = Mid(Form11.Combo1.Text, 1, 2)
StrTemp = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 9 * 4 + Val("&H" & "3F2298")), Hex(Val("&H" & Tilesets) * 9 * 4 + 35 + Val("&H" & "3F2298")))
TileOffset = Mid(StrTemp, 5, 2) & Mid(StrTemp, 3, 2) & Mid(StrTemp, 1, 2)
TileLength2 = Val("&H" & Mid(StrTemp, 11, 2) & Mid(StrTemp, 9, 2))
Form9.Text1.Text = Form9.Text1.Text & "The amount of Tiles is " & str(TileLength2 / 32) & vbCrLf
TextMAPDataOffset = Mid(StrTemp, 45, 2) & Mid(StrTemp, 43, 2) & Mid(StrTemp, 41, 2)
paletteOffset = Mid(StrTemp, 21, 2) & Mid(StrTemp, 19, 2) & Mid(StrTemp, 17, 2)
ReDim Palette256(16, 16)
ReDim Tile88(TileLength2 / 16 + 64)
Form9.Text1.Text = Form9.Text1.Text & "impoting 8 * 8 Tiles Data......" & vbCrLf
Dim TextMapData As String
StrTemp = ReadFileHexWithByteInterchange(gbafilepath, TileOffset, Hex(Val("&H" & TileOffset) + TileLength2))
'-------------------------------------------------------------权宜之策
str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F8098")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F8098")))
'str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F91D8")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F91D8")))             '???
For i = 0 To 15
str2 = Mid(str1, 4 * i + 1, 2)
Mid(str1, 4 * i + 1, 2) = Mid(str1, 4 * i + 3, 2)
Mid(str1, 4 * i + 3, 2) = str2
Next i
For i = 0 To 15
str2 = Mid(str1, 4 * i + 1, 4)
str3 = ReadFileHex(gbafilepath, Hex(Val("&H" & str2) * 8 + Val("&H" & "3F7828")), Hex(Val("&H" & str2) * 8 + Val("&H" & "3F782F")))
str2 = Mid(str3, 9, 8)
str4 = Mid(str3, 1, 2)
str3 = Mid(str3, 5, 2)
If str4 = "03" Or str4 = "06" Then
str3 = Hex(Val("&H" & str3) - 1)
Else
str3 = "00"
End If
str2 = Mid(str2, 7, 2) & Mid(str2, 5, 2) & Mid(str2, 3, 2) & Mid(str2, 1, 2)
str2 = Hex(Val("&H" & str2) - Val("&H8000000") + Val("&H" & str3) * 128)
str2 = ReadFileHexWithByteInterchange(gbafilepath, str2, Hex(Val("&H" & str2) + 127))
DoEvents
For j = 0 To 3
Tile88(i * 4 + j) = Mid(str2, j * 64 + 1, 64)
Next j
Next i
Tile88(64) = Replace(Space(64), Chr(32), "0")
'-------------------------------------------------------------End  权宜之策
For i = 0 To (TileLength2 / 16) - 1
Tile88(i + 65) = Mid(StrTemp, 64 * i + 1, 64)
DoEvents
Next i

Form9.Text1.Text = Form9.Text1.Text & "impoting and making 16 * 16 Tiles Data......" & vbCrLf
TextMapData = ReadFileHex(gbafilepath, TextMAPDataOffset, Hex(Val("&H" & TextMAPDataOffset) + 8192))
For i = 0 To 4095
StrTemp = Mid(TextMapData, 4 * i + 1, 2)
Mid(TextMapData, 4 * i + 1, 2) = Mid(TextMapData, 4 * i + 3, 2)
Mid(TextMapData, 4 * i + 3, 2) = StrTemp
DoEvents
Next i

ReDim Tile16(1024)
Dim r0 As Long, r1 As Long, r2 As Long
For i = 0 To 1023
r0 = i * 4
r2 = r0 Or 1
Tile16(i) = Mid(TextMapData, r0 * 2 * 2 + 1, 4)
r1 = (r2 + 1) * 2 ^ 16
Tile16(i) = Tile16(i) & Mid(TextMapData, r2 * 2 * 2 + 1, 4)
r0 = 128 * 2 ^ 9
r2 = r1 + r0
r1 = RSH(r1, 15)
Tile16(i) = Tile16(i) & Mid(TextMapData, r1 * 2 + 1, 4)
r2 = RSH(r2, 15)
Tile16(i) = Tile16(i) & Mid(TextMapData, r2 * 2 + 1, 4)
DoEvents
Next i

Form9.Text1.Text = Form9.Text1.Text & "impoting palette 256 ......" & vbCrLf
StrTemp = ReadFileHex(gbafilepath, paletteOffset, Hex(Val("&H" & paletteOffset) + 256 * 2 - 1))
For j = 0 To 15
For i = 0 To 15
Palette256(i, j) = RGB555ToRGB888(Mid(StrTemp, 64 * j + 4 * i + 1, 4))
DoEvents
Next i
Next j
Dim a As Boolean
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i, j, Hex(i + 8 * j), Form11.Picture1)
DoEvents
Next i
Next j
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i + 8, j, Hex(128 + i + 8 * j), Form11.Picture1)
DoEvents
Next i
Next j

End Sub

Private Sub Command1_Click()
Dim str As String
Form11.Text1.Text = ""
str = DecompressRLE(Form11.Text2.Text)
If str = "" Then Exit Sub
Dim i As Integer, j As Integer
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
Form11.Text1.Text = Form11.Text1.Text & " " & TextMap(i, j)
Next i
Form11.Text1.Text = Form11.Text1.Text & vbCrLf
Next j
layerHeight = 0
layerWidth = 0
Erase TextMap()
End Sub

Private Sub Form_Load()
Form11.width = 15825
Form11.height = 8565
Form11.Left = Form4.width
Form11.Top = 0
Form11.Combo1.FontSize = 15
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
Form11.Combo1.AddItem "21  Junkyard"
Form11.Combo1.AddItem "22  Junkyard"
Form11.Combo1.AddItem "23  Pinball"
Form11.Combo1.AddItem "24  Pinball"
Form11.Combo1.AddItem "25  Pinball (with Gorilla)"
Form11.Combo1.AddItem "26  Jungle"
Form11.Combo1.AddItem "27  40 Below Fridge?"
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
End Sub

