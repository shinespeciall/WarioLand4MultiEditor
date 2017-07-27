Attribute VB_Name = "Module2"
Public MODfilepath As String
Public IsDeliver As Boolean
Public WholeRoomChange As Boolean
Public ExistUnchangeableLayer0 As Boolean

Public Palette16Color() As Long   'used in form 7
Public Palette256() As Long       'used in form 10

Public Tile16() As String
Public Tile88() As String

Public TileMOD() As String
Public NowTileMOD() As String

Public L0_LB_000() As String
Public L1_LB_000() As String        'three for now
Public L2_LB_000() As String

Public L0_LB_001() As String
Public L1_LB_001() As String        'three for temp and undo
Public L2_LB_001() As String

Public L0_PostCompDataLength As Long
Public L1_PostCompDataLength As Long
Public L2_PostCompDataLength As Long

Public layerPriority() As Integer
Public PostlayerCompDataLength() As Long
Public LastLayerChange As Integer

Public Layer0Height As Integer
Public Layer0Width As Integer

Public DotSize As Integer
Public MODforSave() As String
Public EVA As Integer

Public Function RGB555ToRGB888(ByVal RGB555 As String) As Long
If Len(RGB555) <> 4 Then
RGB555ToRGB888 = 0
Exit Function
End If

Dim str As String
str = Mid$(RGB555, 1, 2)
Mid$(RGB555, 1, 2) = Mid$(RGB555, 3, 2)
Mid$(RGB555, 3, 2) = str

Dim R8 As Long, G8 As Long, B8 As Long
B8 = BIN_to_DEC(Mid$(hextoBin(RGB555), 2, 5) & Mid$(hextoBin(RGB555), 2, 3))
G8 = BIN_to_DEC(Mid$(hextoBin(RGB555), 7, 5) & Mid$(hextoBin(RGB555), 7, 3))
R8 = BIN_to_DEC(Mid$(hextoBin(RGB555), 12, 5) & Mid$(hextoBin(RGB555), 12, 3))

'I think I mistake something here
RGB555ToRGB888 = B8 * 256 * 256 + G8 * 256 + R8
If Mid$(hextoBin(RGB555), 1, 1) = "1" Then
RGB555ToRGB888 = R8 * 256 * 256 + G8 * 256 + B8
Debug.Print "Existing order reverse event in color change!"
End If
End Function

Public Function hextoBin(ByVal X As String) As String
Dim Bina As String, i As Integer
Bina = ""
For i = 1 To Len(X)
Select Case Mid$(X, i, 1)
    Case "0"
        Bina = Bina & "0000"
    Case "1"
        Bina = Bina & "0001"
    Case "2"
        Bina = Bina & "0010"
    Case "3"
        Bina = Bina & "0011"
    Case "4"
        Bina = Bina & "0100"
    Case "5"
        Bina = Bina & "0101"
    Case "6"
        Bina = Bina & "0110"
    Case "7"
        Bina = Bina & "0111"
    Case "8"
        Bina = Bina & "1000"
    Case "9"
        Bina = Bina & "1001"
    Case "A"
        Bina = Bina & "1010"
    Case "B"
        Bina = Bina & "1011"
    Case "C"
        Bina = Bina & "1100"
    Case "D"
        Bina = Bina & "1101"
    Case "E"
        Bina = Bina & "1110"
    Case "F"
        Bina = Bina & "1111"
End Select
Next i
hextoBin = Bina
End Function

Public Function BIN_to_DEC(ByVal Bin As String) As Long
    Dim i As Long, result As Long
    For i = 1 To Len(Bin)
        result = result * 2 + Val(Mid$(Bin, i, 1))
    Next i
    BIN_to_DEC = result
End Function

'Public Function LSH(ByVal X As Long, ByVal B As Integer) As Long
'LSH = BIN_to_DEC(D2B(X) & Replace(Space(B), Chr(32), "0"))
'End Function

Public Function RSH(ByVal X As Long, ByVal b As Integer) As Long
If Len(D2B(X)) <= b Then
RSH = 0
Else
RSH = BIN_to_DEC(Mid$(D2B(X), 1, Len(D2B(X)) - b))
End If
End Function

Public Function D2B(ByVal Dnum As Long) As String
Dim xx As String
Dim yy As Integer
xx = "" '字串累加清空
Do While Dnum > 0 '循环取余至小于1
yy = Dnum Mod 2 '除2取余
Dnum = Dnum \ 2 '除2取整
xx = str(yy) & xx '字串累加
Loop
D2B = Replace(xx, Chr(32), "") '返回字串
End Function

Public Function DrawTile16(ByVal lenpos As Long, ByVal heipos As Long, ByVal TileWord As String, ByVal picbox As PictureBox, Optional Cover As Boolean, Optional SizeOfDot As Integer) As Boolean
On Error Resume Next
'lenpos and heipos are position Index for Tile16
Dim Wrd As String      '处理当前字段
Dim Tile8() As String
Dim i As Integer, j As Integer, strtmp As String
Dim k As Long             '调色板专用

'-------------------------------------first Tile------------------------------
Wrd = Mid$(Tile16(Val("&H" & TileWord)), 1, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(7 - j, i)
Tile8(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(j, 7 - i)
Tile8(j, 7 - i) = strtmp
Next j
Next i
End If
lenpos = lenpos * SizeOfDot * 16          '作为基址
heipos = heipos * SizeOfDot * 16
k = BIN_to_DEC(Mid$(Wrd, 1, 4))
For i = 0 To 7                       '作图
For j = 0 To 7
If Cover = True Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + (SizeOfDot - 1), heipos + i * SizeOfDot + (SizeOfDot - 1)), vbBlack, BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot > 0 Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot < 1 Then picbox.Line (lenpos + j * 24, heipos + i * 24)-(lenpos + j * 24 + 23, heipos + i * 24 + 23), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
Next j
Next i

'-------------------------------------Second Tile------------------------------
Wrd = Mid$(Tile16(Val("&H" & TileWord)), 5, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(7 - j, i)
Tile8(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(j, 7 - i)
Tile8(j, 7 - i) = strtmp
Next j
Next i
End If
lenpos = lenpos + SizeOfDot * 8          '作为基址
heipos = heipos
k = BIN_to_DEC(Mid$(Wrd, 1, 4))
For i = 0 To 7                       '作图
For j = 0 To 7
If Cover = True Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + (SizeOfDot - 1), heipos + i * SizeOfDot + (SizeOfDot - 1)), vbBlack, BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot > 0 Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot < 1 Then picbox.Line (lenpos + j * 24, heipos + i * 24)-(lenpos + j * 24 + 23, heipos + i * 24 + 23), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
Next j
Next i

'-------------------------------------Third Tile------------------------------
Wrd = Mid$(Tile16(Val("&H" & TileWord)), 9, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(7 - j, i)
Tile8(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(j, 7 - i)
Tile8(j, 7 - i) = strtmp
Next j
Next i
End If
lenpos = lenpos - SizeOfDot * 8          '作为基址
heipos = heipos + SizeOfDot * 8
k = BIN_to_DEC(Mid$(Wrd, 1, 4))
For i = 0 To 7                       '作图
For j = 0 To 7
If Cover = True Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + (SizeOfDot - 1), heipos + i * SizeOfDot + (SizeOfDot - 1)), vbBlack, BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot > 0 Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot < 1 Then picbox.Line (lenpos + j * 24, heipos + i * 24)-(lenpos + j * 24 + 23, heipos + i * 24 + 23), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
Next j
Next i

'-------------------------------------Fourth Tile------------------------------
Wrd = Mid$(Tile16(Val("&H" & TileWord)), 13, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(7 - j, i)
Tile8(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8(j, i)
Tile8(j, i) = Tile8(j, 7 - i)
Tile8(j, 7 - i) = strtmp
Next j
Next i
End If
lenpos = lenpos + SizeOfDot * 8          '作为基址
heipos = heipos
k = BIN_to_DEC(Mid$(Wrd, 1, 4))
For i = 0 To 7                       '作图
For j = 0 To 7
If Cover = True Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + (SizeOfDot - 1), heipos + i * SizeOfDot + (SizeOfDot - 1)), vbBlack, BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot > 0 Then picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
If ((Val("&H" & "0" & Tile8(j, i)) <> 0 And Palette256(Val("&H" & "0" & Tile8(j, i)), k) <> Palette256(0, k)) Or Val("&H" & TileWord) = 2) And SizeOfDot > 1 Then picbox.Line (lenpos + j * 24, heipos + i * 24)-(lenpos + j * 24 + 23, heipos + i * 24 + 23), Palette256(Val("&H" & "0" & Tile8(j, i)), k), BF
Next j
Next i

Erase Tile8()
DrawTile16 = True
End Function

Public Function Min(a As Single, b As Single) As Single
If a <= b Then Min = a Else Min = b
End Function

Public Function GetAlphaBlendColor(ByVal TopColor As Long, BottomColor As Long, RenderEVA As Integer) As Long    'Some glitches have been found and haven't been fix yet
Dim rRed As Long, rGreen As Long, rBlue As Long         'the order in VB6 is &BBGGRR
If BottomColor = 0 Then
GetAlphaBlendColor = TopColor
Exit Function
End If
rRed = Min(((TopColor And CLng("&HFF")) * RenderEVA) \ 16 + (BottomColor And CLng("&HFF")), 255)
'rRed = Min(((TopColor And CLng("&HFF")) * (16 - RenderEVA)) \ 16 + (BottomColor And CLng("&HFF")), 255)
rGreen = Min((((TopColor And CLng("&HFF00")) / 256) * RenderEVA) \ 16 + ((BottomColor And CLng("&HFF00")) / 256), 255)
'rGreen = Min((((TopColor And CLng("&HFF00")) / 256) * (16 - RenderEVA)) \ 16 + ((BottomColor And CLng("&HFF00")) / 256), 255)
rBlue = Min((((TopColor And CLng("&HFF0000")) / 65536) * RenderEVA) \ 16 + ((BottomColor And CLng("&HFF0000")) / 65536), 255)
'rBlue = Min((((TopColor And CLng("&HFF0000")) / 65536) * (16 - RenderEVA)) \ 16 + ((BottomColor And CLng("&HFF0000")) / 65536), 255)
GetAlphaBlendColor = rBlue * 65536 + rGreen * 256 + rRed
End Function

Public Function DrawTile16_Alpha(ByVal lenpos As Long, ByVal heipos As Long, ByVal TopTileWord As String, ByVal MiddleTileWord As String, ByVal BottomTileWord As String, ByVal picbox As PictureBox, ByVal EVALng As Integer, ByVal SizeOfDot As Integer, Optional Cover As Boolean) As Boolean
On Error Resume Next    'Some glitches have been found and haven't been fix yet
If SizeOfDot < 1 Then
DrawTile16_Alpha = False
Exit Function
End If
'lenpos and heipos are position Index for Tile16
Dim Wrd As String      '处理当前字段
Dim Tile8_T() As String, Tile8_M() As String, Tile8_B() As String
Dim i As Integer, j As Integer, strtmp As String
Dim kT As Long, kM As Long, kB As Long             '调色板专用
Dim RealBottomColor As Long, RealTopColor As Long
'-------------------------------------first Tile------------------------------
'-------------Top-------------
Wrd = Mid$(Tile16(Val("&H" & TopTileWord)), 1, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_T(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_T(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(7 - j, i)
Tile8_T(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(j, 7 - i)
Tile8_T(j, 7 - i) = strtmp
Next j
Next i
End If
kT = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Middle-------------
Wrd = Mid$(Tile16(Val("&H" & MiddleTileWord)), 1, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_M(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_M(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(7 - j, i)
Tile8_M(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(j, 7 - i)
Tile8_M(j, 7 - i) = strtmp
Next j
Next i
End If
kM = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Bottom-------------
Wrd = Mid$(Tile16(Val("&H" & BottomTileWord)), 1, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_B(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_B(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(7 - j, i)
Tile8_B(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(j, 7 - i)
Tile8_B(j, 7 - i) = strtmp
Next j
Next i
End If
kB = BIN_to_DEC(Mid$(Wrd, 1, 4))
'------End of Load Tile Data-----
lenpos = lenpos * SizeOfDot * 16          '作为基址
heipos = heipos * SizeOfDot * 16
If Cover = True Then picbox.Line (lenpos, heipos)-(lenpos + 7 * SizeOfDot + SizeOfDot - 1, heipos + 7 * SizeOfDot + SizeOfDot - 1), vbBlack, BF
For i = 0 To 7                       '作图
For j = 0 To 7
If Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM) <> Palette256(0, kM) Then
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM)
Else
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_B(j, i)), kB)
End If
If RealBottomColor = Palette256(0, kB) Then RealBottomColor = 0
RealTopColor = Palette256(Val("&H" & "0" & Tile8_T(j, i)), kT)
If RealTopColor = Palette256(0, kT) Then RealTopColor = 0
RealBottomColor = GetAlphaBlendColor(RealTopColor, RealBottomColor, EVALng)
picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), RealBottomColor, BF
Next j
Next i

'-------------------------------------Second Tile------------------------------
'-------------Top-------------
Wrd = Mid$(Tile16(Val("&H" & TopTileWord)), 5, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_T(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_T(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(7 - j, i)
Tile8_T(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(j, 7 - i)
Tile8_T(j, 7 - i) = strtmp
Next j
Next i
End If
kT = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Middle-------------
Wrd = Mid$(Tile16(Val("&H" & MiddleTileWord)), 5, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_M(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_M(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(7 - j, i)
Tile8_M(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(j, 7 - i)
Tile8_M(j, 7 - i) = strtmp
Next j
Next i
End If
kM = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Bottom-------------
Wrd = Mid$(Tile16(Val("&H" & BottomTileWord)), 5, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_B(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_B(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(7 - j, i)
Tile8_B(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(j, 7 - i)
Tile8_B(j, 7 - i) = strtmp
Next j
Next i
End If
kB = BIN_to_DEC(Mid$(Wrd, 1, 4))
'------End of Load Tile Data-----
lenpos = lenpos + SizeOfDot * 8          '作为基址
heipos = heipos
If Cover = True Then picbox.Line (lenpos, heipos)-(lenpos + 7 * SizeOfDot + SizeOfDot - 1, heipos + 7 * SizeOfDot + SizeOfDot - 1), vbBlack, BF
For i = 0 To 7                       '作图
For j = 0 To 7
If Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM) <> Palette256(0, kM) Then
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM)
Else
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_B(j, i)), kB)
End If
If RealBottomColor = Palette256(0, kB) Then RealBottomColor = 0
RealTopColor = Palette256(Val("&H" & "0" & Tile8_T(j, i)), kT)
If RealTopColor = Palette256(0, kT) Then RealTopColor = 0
RealBottomColor = GetAlphaBlendColor(RealTopColor, RealBottomColor, EVALng)
picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), RealBottomColor, BF
Next j
Next i

'-------------------------------------Third Tile------------------------------
'-------------Top-------------
Wrd = Mid$(Tile16(Val("&H" & TopTileWord)), 9, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_T(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_T(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(7 - j, i)
Tile8_T(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(j, 7 - i)
Tile8_T(j, 7 - i) = strtmp
Next j
Next i
End If
kT = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Middle-------------
Wrd = Mid$(Tile16(Val("&H" & MiddleTileWord)), 9, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_M(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_M(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(7 - j, i)
Tile8_M(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(j, 7 - i)
Tile8_M(j, 7 - i) = strtmp
Next j
Next i
End If
kM = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Bottom-------------
Wrd = Mid$(Tile16(Val("&H" & BottomTileWord)), 9, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_B(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_B(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(7 - j, i)
Tile8_B(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(j, 7 - i)
Tile8_B(j, 7 - i) = strtmp
Next j
Next i
End If
kB = BIN_to_DEC(Mid$(Wrd, 1, 4))
'------End of Load Tile Data-----
lenpos = lenpos - SizeOfDot * 8          '作为基址
heipos = heipos + SizeOfDot * 8
If Cover = True Then picbox.Line (lenpos, heipos)-(lenpos + 7 * SizeOfDot + SizeOfDot - 1, heipos + 7 * SizeOfDot + SizeOfDot - 1), vbBlack, BF
For i = 0 To 7                       '作图
For j = 0 To 7
If Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM) <> Palette256(0, kM) Then
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM)
Else
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_B(j, i)), kB)
End If
If RealBottomColor = Palette256(0, kB) Then RealBottomColor = 0
RealTopColor = Palette256(Val("&H" & "0" & Tile8_T(j, i)), kT)
If RealTopColor = Palette256(0, kT) Then RealTopColor = 0
RealBottomColor = GetAlphaBlendColor(RealTopColor, RealBottomColor, EVALng)
picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), RealBottomColor, BF
Next j
Next i

'-------------------------------------Fourth Tile------------------------------
'-------------Top-------------
Wrd = Mid$(Tile16(Val("&H" & TopTileWord)), 13, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_T(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_T(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(7 - j, i)
Tile8_T(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_T(j, i)
Tile8_T(j, i) = Tile8_T(j, 7 - i)
Tile8_T(j, 7 - i) = strtmp
Next j
Next i
End If
kT = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Middle-------------
Wrd = Mid$(Tile16(Val("&H" & MiddleTileWord)), 13, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_M(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_M(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(7 - j, i)
Tile8_M(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_M(j, i)
Tile8_M(j, i) = Tile8_M(j, 7 - i)
Tile8_M(j, 7 - i) = strtmp
Next j
Next i
End If
kM = BIN_to_DEC(Mid$(Wrd, 1, 4))
'-------------Bottom-------------
Wrd = Mid$(Tile16(Val("&H" & BottomTileWord)), 13, 4)
Wrd = Right("0000000000000000" & hextoBin(Wrd), 16)
ReDim Tile8_B(8, 8)
strtmp = Tile88(BIN_to_DEC(Mid$(Wrd, 7, 10)))
For i = 0 To 7
For j = 0 To 7
Tile8_B(j, i) = Mid$(strtmp, 1 + j + 8 * i, 1)
Next j
Next i
If Mid$(Wrd, 6, 1) = 1 Then      '水平翻转
For i = 0 To 7
For j = 0 To 3
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(7 - j, i)
Tile8_B(7 - j, i) = strtmp
Next j
Next i
End If
If Mid$(Wrd, 5, 1) = 1 Then      '垂直翻转
For i = 0 To 3
For j = 0 To 7
strtmp = Tile8_B(j, i)
Tile8_B(j, i) = Tile8_B(j, 7 - i)
Tile8_B(j, 7 - i) = strtmp
Next j
Next i
End If
kB = BIN_to_DEC(Mid$(Wrd, 1, 4))
'------End of Load Tile Data-----
lenpos = lenpos + SizeOfDot * 8          '作为基址
heipos = heipos
If Cover = True Then picbox.Line (lenpos, heipos)-(lenpos + 7 * SizeOfDot + SizeOfDot - 1, heipos + 7 * SizeOfDot + SizeOfDot - 1), vbBlack, BF
For i = 0 To 7                       '作图
For j = 0 To 7
If Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM) <> Palette256(0, kM) Then
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_M(j, i)), kM)
Else
RealBottomColor = Palette256(Val("&H" & "0" & Tile8_B(j, i)), kB)
End If
If RealBottomColor = Palette256(0, kB) Then RealBottomColor = 0
RealTopColor = Palette256(Val("&H" & "0" & Tile8_T(j, i)), kT)
If RealTopColor = Palette256(0, kT) Then RealTopColor = 0
RealBottomColor = GetAlphaBlendColor(RealTopColor, RealBottomColor, EVALng)
picbox.Line (lenpos + j * SizeOfDot, heipos + i * SizeOfDot)-(lenpos + j * SizeOfDot + SizeOfDot - 1, heipos + i * SizeOfDot + SizeOfDot - 1), RealBottomColor, BF
Next j
Next i

Erase Tile8_B()
Erase Tile8_M()
Erase Tile8_T()

DrawTile16_Alpha = True
End Function
