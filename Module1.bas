Attribute VB_Name = "Module1"
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////This part is for control use

Public gbafilepath As String         'save gba file path and name

Public Hexstream1 As String         'load current layer1 decompressed room all Hex stream
Public Hexstream2 As String         'load current layer2 decompressed room all Hex stream

Public widtha1 As String        'save layer's width
Public heighta2 As String       'save layer's height
Public transmita3 As String      'transmit a3

Public leftzerozero1 As Long         'save the "00" data number

Public BeforeLine As Integer           'for Form1 drawing line, rectangle and print font

Public layer1compressdatalength As Long    'store layer compress data length, 单位是4个bit，半个字节
Public layer2compressdatalength As Long

Public startoffset As String     'Just store in Hex, if use, we can change it to Dec.
Public PointerOffset1 As String  'make index in case of expand other pointer Offset varients

Public IfisNewRoom As Boolean    'decide if form2 show to create a new room
Public IfisNewRoomConnectionDataBuffer As Boolean
Public RoomConnectionDataBuffer As String

Public LevelStartStream As String
Public LevelStartStreamOffset As String
Public LevelNumber As String     'store level number which can be got from 030000023 h
Public LevelRoomIndex As String                          'count from 1
Public LevelAllRoomPointerandDataBaseOffset As String
Public LevelAllRoomPointerandDataallHex As String
Public RoomElementOffset As String
Public LevelChangeRoomStreamOffset As String
Public LevelChangeRoomStreamPointerOffset As String

Public SaveDatabuffer() As String
Public SaveDataOffset() As String

Public TempPointerValue() As String

Public RoomElementFirstOffset As String

'******************************************************************************from Form 6 for global use
Public CameraCotrolString As String
Public CameraCotrolPointerOffset As String      '存放（指向指针表表头位置的指针）的地址
Public RoomCameraStringPointerOffset As String     '存放（指向Room的Camera控制流字符串的指针）的地址
Public LengthOfAllPointer As Long               '指针表总长，单位是Byte
'******************************************************************************
Public WasCameraControlStringChange As Boolean
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Function CompressDataOnly(ByVal Hexstream As String) As String   'compress data only, and value "FF" has not been check
Dim OutputStream As String
Dim str1 As String, str2 As String    'store now text in byte

Dim Num1 As Integer    'count for non repeat byte
Dim num2 As Integer    'count for  repeat byte

Dim shiftoffset As Long   '单位是8个bit，半个字节
Dim tempstream As String  '缓存非重复字节数据

Num1 = 0
num2 = Val("&H" & "80")
shiftoffset = -1

Do
DoEvents
shiftoffset = shiftoffset + 2
str1 = Mid(Hexstream, shiftoffset, 2)
str2 = Mid(Hexstream, shiftoffset + 2, 2)

If str2 <> str1 Then
Num1 = Num1 + 1

    If num2 = Val("&H" & "80") And Num1 < Val("&H" & "7E") And (shiftoffset < Len(Hexstream) - 3) Then    'less then 7Eh
    tempstream = tempstream & str1
    ElseIf num2 = Val("&H" & "80") And Num1 = Val("&H" & "7E") Then               'now equal to 7Eh
    OutputStream = OutputStream & "7F" & tempstream & str1 & str2
    Num1 = 0
    shiftoffset = shiftoffset + 2
    tempstream = ""
    ElseIf num2 = Val("&H" & "80") And Num1 < Val("&H" & "7E") And (shiftoffset = Len(Hexstream) - 3) Then               'now equal to 7Eh and to the end
    OutputStream = OutputStream & Right("00" & Hex(Num1 + 1), 2) & tempstream & str1 & str2
    ElseIf num2 > Val("&H" & "80") And num2 < Val("&H" & "FF") And (shiftoffset < Len(Hexstream) - 3) And Num1 = 1 Then
    OutputStream = OutputStream & Right("00" & Hex(num2 + 1), 2) & str1
    num2 = Val("&H" & "80")
    Num1 = 0
    ElseIf num2 > Val("&H" & "80") And (shiftoffset = Len(Hexstream) - 3) Then
    OutputStream = OutputStream & Right("00" & Hex(num2 + 2), 2) & str1 & "01" & str2
    End If
ElseIf str1 = str2 Then

    If Num1 > 0 And (shiftoffset < Len(Hexstream) - 3) Then   '出现了一串不重复字符后面的两个重复字符
    OutputStream = OutputStream & Right("00" & Hex(Num1), 2) & tempstream
    Num1 = 0
    num2 = Val("&H" & "81")
    tempstream = ""
    ElseIf Num1 > 0 And (shiftoffset = Len(Hexstream) - 3) And Num1 < Val("&H" & "7E") Then   'to the end   我觉得这几句会有错，但是因为最后一行一般都是清一色的40或00，所以这几句一般用不到
    OutputStream = OutputStream & Right("00" & Hex(Num1 + 2), 2) & tempstream & str1 & str2
    tempstream = ""
    ElseIf Num1 > 0 And (shiftoffset = Len(Hexstream) - 3) And Num1 = Val("&H" & "7E") Then   'to the end   我觉得这几句会有错，但是因为最后一行一般都是清一色的40或00，所以这几句一般用不到
    OutputStream = OutputStream & Right("00" & Hex(Num1 + 1), 2) & tempstream & str1 & "01" & str2
    tempstream = ""
    ElseIf Num1 > 0 And (shiftoffset = Len(Hexstream) - 3) And Num1 = Val("&H" & "7F") Then   'to the end   我觉得这几句会有错，但是因为最后一行一般都是清一色的40或00，所以这几句一般用不到
    OutputStream = OutputStream & Right("00" & Hex(Num1), 2) & tempstream & "82" & str2
    tempstream = ""
    ElseIf Num1 = 0 And num2 < Val("&H" & "FE") And (shiftoffset < Len(Hexstream) - 3) Then
    num2 = num2 + 1
    ElseIf Num1 = 0 And num2 < Val("&H" & "FE") And (shiftoffset = Len(Hexstream) - 3) Then
    num2 = num2 + 1
    OutputStream = OutputStream & Right("00" & Hex(num2 + 1), 2) & str1
    ElseIf Num1 = 0 And num2 = Val("&H" & "FE") Then
    OutputStream = OutputStream & "FF" & str1
    num2 = Val("&H" & "80")
    End If
End If

DoEvents

If shiftoffset = Len(Hexstream) - 3 Then
Exit Do
End If

Form2.Label1.Caption = "Output:" & str(shiftoffset) & "/" & str(Len(Hexstream) - 2)
Loop

CompressDataOnly = OutputStream
End Function

Public Function FindSpace(ByVal filepath As String, ByVal StartOffset1 As String, ByVal EndOffset1 As String, ByVal SpaceStr As String, ByVal SpaceLen As Long) As String
If StartOffset1 = "" Then StartOffset1 = "00"
If filepath = "" Then
FindSpace = ""
MsgBox "没有制定ROM File！", vbOKOnly + vbExclamation, "Warning!"
Exit Function
End If
If Val("&H" & EndOffset1) - Val("&H" & StartOffset1) + 1 < SpaceLen Then '
FindSpace = "FFFFFFFF"
Exit Function
End If
If SpaceStr = "" Then SpaceStr = "00"
Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, Val("&H" & StartOffset1) + 1, ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim j As Long         '计数器
For i = LBound(ROMallbyte) To LBound(ROMallbyte) + Val("&H" & EndOffset1) - Val("&H" & StartOffset1)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Next i
Erase ROMallbyte()
For i = 0 To Val("&H" & EndOffset1) - Val("&H" & StartOffset1)
'Form2.Label8.Caption = "搜索可能的源地址中的Free Space，进度：" & CStr(i) & CStr(Val("&H" & EndOffset1) - Val("&H" & StartOffset1))
If Mid(ROMallHex, 2 * i + 1, 2) = SpaceStr Then j = j + 1
If Mid(ROMallHex, 2 * i + 1, 2) <> SpaceStr Then j = 0
If Val("&H" & StartOffset1) + i > Val("&H" & EndOffset1) - SpaceLen Then
FindSpace = "FFFFFFFF"                                                   '返回错误代码
Exit Function
End If
If j = SpaceLen Then
FindSpace = Hex(Val("&H" & StartOffset1) + i - j + 1)
Exit For
End If
DoEvents
Next i
End Function

Public Function ReadFileHex(ByVal filepath As String, ByVal StartOffset2 As String, ByVal EndOffset2 As String) As String
If StartOffset2 = "" Then StartOffset2 = "00"
Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
If EndOffset2 = "" Or Val("&H" & EndOffset2) = 0 Then EndOffset2 = Hex(LOF(1) - 1)
Get #1, Val("&H" & StartOffset2) + 1, ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim j As Long         '计数器
For i = LBound(ROMallbyte) To LBound(ROMallbyte) + (Val("&H" & EndOffset2) - Val("&H" & StartOffset2))
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Next i
Erase ROMallbyte()
ReadFileHex = ROMallHex
End Function

Public Function ReadFileHexWithByteInterchange(ByVal filepath As String, ByVal StartOffset2 As String, ByVal EndOffset2 As String) As String
If StartOffset2 = "" Then StartOffset2 = "00"
Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
If EndOffset2 = "" Or Val("&H" & EndOffset2) = 0 Then EndOffset2 = Hex(LOF(1) - 1)
Get #1, Val("&H" & StartOffset2) + 1, ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim j As Long         '计数器
Dim n1 As String, n2 As String
For i = LBound(ROMallbyte) To LBound(ROMallbyte) + (Val("&H" & EndOffset2) - Val("&H" & StartOffset2))
ROMallHex = ROMallHex & Hex(ROMallbyte(i) And 15) & Mid(Hex(ROMallbyte(i) And 240), 1, 1)
DoEvents
Next i
Erase ROMallbyte()
ReadFileHexWithByteInterchange = ROMallHex
End Function

Public Function strcmp(str1 As String, str2 As String) As Integer
If (Len(str1) < Len(str2)) Or (Len(str1) > Len(str2)) Then
strcmp = -1
Exit Function
End If
Dim i As Long
For i = 1 To Len(str1)
If Mid(str1, i, 1) <> Mid(str2, i, 1) Then
strcmp = i
Exit Function
End If
Next i
strcmp = 0
End Function

Public Function SaveCameraString(StrTemp As String) As Boolean         'not support resave
If SaveDataOffset(95) <> "" Then
    SaveCameraString = False
    Exit Function
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 24 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
Next i
SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = "03"

i = i + 1
Dim TempPointer As String

StrTemp = Replace(StrTemp, Chr(32), "")
StrTemp = Replace(StrTemp, Chr(13), "")
StrTemp = Replace(StrTemp, Chr(10), "")

If RoomCameraStringPointerOffset = "" Then               '以前不存在Camera控制
        SaveDataOffset(i) = SaveDatabuffer(0)        '先写新的Camera控制流数据
        TempPointer = Right("00" & Hex(Val("&H" & SaveDataOffset(i)) + Val("&H8000000")), 8)
        TempPointer = Mid(TempPointer, 7, 2) & Mid(TempPointer, 5, 2) & Mid(TempPointer, 3, 2) & Mid(TempPointer, 1, 2)
        SaveDatabuffer(i) = StrTemp
        SaveDatabuffer(0) = Hex(Val("&H" & SaveDatabuffer(0)) + Len(StrTemp))   '基址重整
        SaveDatabuffer(0) = (SaveDatabuffer(0) \ 4) * 4 + 4
        SaveDataOffset(i + 1) = CameraCotrolPointerOffset      '修改指针表表头指针，接下来计算指针表新位置和长度
        SaveDatabuffer(i + 1) = Right("0000" & Hex(Val("&H" & SaveDatabuffer(0)) + Val("&H8000000")), 8)
        SaveDatabuffer(i + 1) = Mid(SaveDatabuffer(i + 1), 7, 2) & Mid(SaveDatabuffer(i + 1), 5, 2) & Mid(SaveDatabuffer(i + 1), 3, 2) & Mid(SaveDatabuffer(i + 1), 1, 2)    '重置指针，定位了新的指针表地址
        SaveDataOffset(i + 2) = SaveDatabuffer(0)      '写新的指针表
        
        SaveDatabuffer(i + 2) = TempPointer & ReadFileHex(gbafilepath, CameraCotrolPointerOffset, Hex(Val("&H" & CameraCotrolPointerOffset) + LengthOfAllPointer - 1))
        SaveDatabuffer(0) = Hex(Val("&H" & SaveDatabuffer(0)) + LengthOfAllPointer + 4) '基址重整
Else
        If Len(StrTemp) > Len(CameraCotrolString) Then         '以前存在只是现在的比较长
        SaveDataOffset(i) = RoomCameraStringPointerOffset
        TempPointer = Right("0000" & Hex(Val("&H" & SaveDatabuffer(0)) + Val("&H8000000")), 8)
        TempPointer = Mid(TempPointer, 7, 2) & Mid(TempPointer, 5, 2) & Mid(TempPointer, 3, 2) & Mid(TempPointer, 1, 2)
        SaveDatabuffer(i) = TempPointer
        SaveDataOffset(i + 1) = SaveDatabuffer(0)
        SaveDatabuffer(i + 1) = StrTemp
        SaveDatabuffer(0) = Hex(Val("&H" & SaveDatabuffer(0)) + Len(StrTemp))   '基址重整
        Else
        SaveDataOffset(i) = RoomCameraStringPointerOffset
        SaveDatabuffer(i) = StrTemp & Replace(Space(Len(CameraCotrolString) - Len(StrTemp)), Chr(32), "0")
        End If
End If
SaveCameraString = True
End Function

Public Function DEX_to_letter(ByVal Dex As Long) As String
Dim str1 As String
       Select Case Dex
           Case 0: str1 = "0"
           Case 1: str1 = "1"
           Case 2: str1 = "2"
           Case 3: str1 = "3"
           Case 4: str1 = "4"
           Case 5: str1 = "5"
           Case 6: str1 = "6"
           Case 7: str1 = "7"
           Case 8: str1 = "8"
           Case 9: str1 = "9"
           Case 10: str1 = "A"
           Case 11: str1 = "B"
           Case 12: str1 = "C"
           Case 13: str1 = "D"
           Case 14: str1 = "E"
           Case 15: str1 = "F"
           Case 16: str1 = "G"
           Case 17: str1 = "H"
           Case 18: str1 = "I"
           Case 19: str1 = "J"
           Case 20: str1 = "K"
           Case 21: str1 = "L"
           Case 22: str1 = "M"
           Case 23: str1 = "N"
           Case 24: str1 = "O"
           Case 25: str1 = "P"
           Case 26: str1 = "Q"
           Case 27: str1 = "R"
           Case 28: str1 = "S"
           Case 29: str1 = "T"
           Case 30: str1 = "U"
           Case 31: str1 = "V"
           Case 32: str1 = "W"
           Case 33: str1 = "X"
           Case 34: str1 = "Y"
           Case 35: str1 = "Z"
           Case 36: str1 = "a"
           Case 37: str1 = "b"
           Case 38: str1 = "c"
           Case 39: str1 = "d"
           Case 40: str1 = "e"
           Case 41: str1 = "f"
           Case 42: str1 = "g"
           Case 43: str1 = "h"
           Case 44: str1 = "i"
           Case 45: str1 = "j"
           Case 46: str1 = "k"
           Case 47: str1 = "l"
           Case 48: str1 = "m"
           Case 49: str1 = "n"
           Case 50: str1 = "o"
           Case 51: str1 = "p"
           Case 52: str1 = "q"
           Case 53: str1 = "r"
           Case 54: str1 = "s"
           Case 55: str1 = "t"
           Case 56: str1 = "u"
           Case 57: str1 = "v"
           Case 58: str1 = "w"
           Case 59: str1 = "x"
           Case 60: str1 = "y"
           Case 61: str1 = "z"
'*****************************************      Extra
           Case 62: str1 = "."
           Case 63: str1 = "&"
           Case 64: str1 = "'"
           Case 226: str1 = ","
           Case 227: str1 = "。"
           Case 228: str1 = "-"
           Case 229: str1 = "~"
           Case 230: str1 = "…"
           Case 231: str1 = "!"
           Case 232: str1 = "?"
           Case 233: str1 = "("
           Case 234: str1 = ")"
           Case 235: str1 = "┌"
           Case 236: str1 = "┘"
           Case 237: str1 = "╔"
           Case 238: str1 = "╝"
           Case 239: str1 = "["
           Case 240: str1 = "]"
           Case 241: str1 = "℃"
           Case 242: str1 = "-"
'******************************************
           Case 255: str1 = " "
           Case Else: str1 = " "
           End Select
   DEX_to_letter = str1
End Function

Public Function letter_to_DEX(ByVal letter As String) As Long
Dim str1 As Long
       Select Case letter
           Case "0": str1 = 0
           Case "1": str1 = 1
           Case "2": str1 = 2
           Case "3": str1 = 3
           Case "4": str1 = 4
           Case "5": str1 = 5
           Case "6": str1 = 6
           Case "7": str1 = 7
           Case "8": str1 = 8
           Case "9": str1 = 9
           Case "A": str1 = 10
           Case "B": str1 = 11
           Case "C": str1 = 12
           Case "D": str1 = 13
           Case "E": str1 = 14
           Case "F": str1 = 15
           Case "G": str1 = 16
           Case "H": str1 = 17
           Case "I": str1 = 18
           Case "J": str1 = 19
           Case "K": str1 = 20
           Case "L": str1 = 21
           Case "M": str1 = 22
           Case "N": str1 = 23
           Case "O": str1 = 24
           Case "P": str1 = 25
           Case "Q": str1 = 26
           Case "R": str1 = 27
           Case "S": str1 = 28
           Case "T": str1 = 29
           Case "U": str1 = 30
           Case "V": str1 = 31
           Case "W": str1 = 32
           Case "X": str1 = 33
           Case "Y": str1 = 34
           Case "Z": str1 = 35
           Case "a": str1 = 36
           Case "b": str1 = 37
           Case "c": str1 = 38
           Case "d": str1 = 39
           Case "e": str1 = 40
           Case "f": str1 = 41
           Case "g": str1 = 42
           Case "h": str1 = 43
           Case "i": str1 = 44
           Case "j": str1 = 45
           Case "k": str1 = 46
           Case "l": str1 = 47
           Case "m": str1 = 48
           Case "n": str1 = 49
           Case "o": str1 = 50
           Case "p": str1 = 51
           Case "q": str1 = 52
           Case "r": str1 = 53
           Case "s": str1 = 54
           Case "t": str1 = 55
           Case "u": str1 = 56
           Case "v": str1 = 57
           Case "w": str1 = 58
           Case "x": str1 = 59
           Case "y": str1 = 60
           Case "z": str1 = 61
'********************************************* Extra
           Case ".": str1 = 62
           Case "&": str1 = 63
           Case "'": str1 = 225
           Case ",": str1 = 226
           Case "。": str1 = 227
           Case "-": str1 = 228
           Case "~": str1 = 229
           Case "…": str1 = 230
           Case "!": str1 = 231
           Case "?": str1 = 232
           Case "(": str1 = 233
           Case ")": str1 = 234
           Case "┌": str1 = 235
           Case "┘": str1 = 236
           Case "╔": str1 = 237
           Case "╝": str1 = 238
           Case "[": str1 = 239
           Case "]": str1 = 240
           Case "℃": str1 = 241
           Case "-": str1 = 242
'*********************************************
           Case " ": str1 = 255
           Case Else: str1 = 255
           End Select
   letter_to_DEX = str1
End Function

Public Function GetLevelNamePointer(level As Long) As String
Dim r3 As String, X As Integer
r3 = "63A3AC"
Select Case level
    Case 0: X = 0
    Case 1: X = 6
    Case 2: X = 7
    Case 3: X = 8
    Case 4: X = 9
    Case 5: X = 12
    Case 6: X = 13
    Case 7: X = 15
    Case 8: X = 14
    Case 9: X = 18
    Case 10: X = 19
    Case 11: X = 20
    Case 12: X = 21
    Case 13: X = 24
    Case 14: X = 25
    Case 15: X = 27
    Case 16: X = 26
    Case 17: X = 10
    Case 18: X = 16
    Case 19: X = 22
    Case 20: X = 28
    Case 21: X = 34
    Case 22: X = 4
    Case 23: X = 30
End Select
r3 = Hex(Val("&H" & r3) + 4 * X)
GetLevelNamePointer = ReadFileHex(gbafilepath, r3, Hex(Val("&H" & r3) + 3))
End Function

Public Function SaveRoomCompData(Offset_ofPostDataPointer As String, compData As String, PostDataLength As Long) As Boolean

Dim i As Integer, j As Long
Dim TempAddress As Long
Dim returnstr As String
Dim str1 As String

For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
Next i
If SaveDataOffset(95) <> "" Then
SaveRoomCompData = False
Exit Function
End If

TempAddress = CLng("&H" & SaveDatabuffer(0))

    returnstr = FindSpace(gbafilepath, "598EEC", "59F291", "00", Len(compData) / 2 + 8)
    If returnstr = "FFFFFFFF" Then
    returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(compData) / 2 + 8)
    End If
    '------------------------出问题的地址可以在此处设置
    'If CLng("&H" & returnstr) >= CLng("&H59AD20") And CLng("&H" & returnstr) <= CLng("&H59AE63") Then
    'returnstr = FindSpace(gbafilepath, "59AE63", "59F291", "00", Len(compData)/ 2 + 6)
    'If returnstr = "FFFFFFFF" Then
    'returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(compData)/ 2 + 8)
    'End If
    'End If
    '-------------------------------------------------
If returnstr = "FFFFFFFF" Then
    SaveDataOffset(i) = SaveDatabuffer(0)
    SaveDatabuffer(i) = compData
    SaveDataOffset(i + 1) = Offset_ofPostDataPointer
    TempAddress = CLng("&H" & "8000000") + CLng("&H" & SaveDataOffset(i))
    SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    TempAddress = CLng("&H" & ReadFileHex(gbafilepath, Offset_ofPostDataPointer, Hex(CLng("&H" & Offset_ofPostDataPointer) + 3)))
    TempAddress = TempAddress - 8
    If PostDataLength > 0 Then
    SaveDataOffset(i + 2) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    For j = 1 To PostDataLength
    str1 = str1 + "00"
    Next j
    SaveDatabuffer(i + 2) = str1
    End If
    SaveDatabuffer(0) = Hex(CLng("&H" & SaveDatabuffer(0)) + Len(compData) / 2)
Else
    SaveDataOffset(i) = Hex(CLng("&H" & returnstr) + 4)
    SaveDatabuffer(i) = compData
    SaveDataOffset(i + 1) = Offset_ofPostDataPointer
    TempAddress = CLng("&H" & "8000000") + CLng("&H" & SaveDataOffset(i))
    SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    TempAddress = CLng("&H" & ReadFileHex(gbafilepath, Offset_ofPostDataPointer, Hex(CLng("&H" & Offset_ofPostDataPointer) + 3)))
    TempAddress = TempAddress - 8
    If PostDataLength > 0 Then
    SaveDataOffset(i + 2) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    For j = 1 To PostDataLength
    str1 = str1 + "00"
    Next j
    SaveDatabuffer(i + 2) = str1
    End If
End If
SaveRoomCompData = True
End Function
