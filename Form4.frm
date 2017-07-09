VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level guide"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   4560
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Room"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      ToolTipText     =   "BG layer will be thrown"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add a room"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.ListBox List6 
      Height          =   1680
      ItemData        =   "Form4.frx":0000
      Left            =   1200
      List            =   "Form4.frx":0002
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   1680
      ItemData        =   "Form4.frx":0004
      Left            =   120
      List            =   "Form4.frx":0006
      TabIndex        =   12
      ToolTipText     =   "double click to change value"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox List4 
      Height          =   1680
      ItemData        =   "Form4.frx":0008
      Left            =   3360
      List            =   "Form4.frx":000A
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   1680
      ItemData        =   "Form4.frx":000C
      Left            =   2280
      List            =   "Form4.frx":000E
      TabIndex        =   10
      ToolTipText     =   "double click to change value"
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   1680
      ItemData        =   "Form4.frx":0010
      Left            =   1200
      List            =   "Form4.frx":0012
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   6840
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "transfer"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1680
      ItemData        =   "Form4.frx":0014
      Left            =   120
      List            =   "Form4.frx":0016
      TabIndex        =   3
      ToolTipText     =   "double click to change value"
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form4.frx":0018
      Left            =   240
      List            =   "Form4.frx":001A
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "layer 0   pointer offset"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Output："
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "layer 1    pointeroffset layer 2   pointeroffset"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "info："
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Form4TextBox2Temp As String
Public RoomNumber As String       '第一个Room记为1

Private Sub Combo1_Change()
Form4.Combo1.Text = "00"
LevelNumber = Mid(Form4.Combo1.Text, 1, 2)
End Sub

Private Sub Combo1_Click()
If gbafilepath = "" Then Exit Sub
Form4.Combo1.Enabled = False
LevelNumber = Mid(Form4.Combo1.Text, 1, 2)
Form4.List1.Clear
Form4.List2.Clear
Form4.List3.Clear
Form4.List4.Clear
Form4.List5.Clear
Form4.List6.Clear
Form4.Text1.Text = ""
Form4.Text2.Text = ""

If gbafilepath = "" Then Exit Sub
    ': 正在搜索……"

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String

Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1

Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
bytenum = 720

Dim offset_639068 As String
offset_639068 = "639068"

For i = LBound(ROMallbyte) + CLng(Val("&H" & offset_639068)) To LBound(ROMallbyte) + CLng(Val("&H" & offset_639068)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Form4.Label2.Caption = "process: Load the data......" & i - LBound(ROMallbyte) - CLng(Val("&H" & offset_639068)) & "/" & bytenum
Next i

'***********************************************************************************************************************改为使用搜索法搜索00 -10, 17
Dim LevelIndex As String
For i = 0 To 30
    LevelIndex = Mid(ROMallHex, i * 24 + 1, 2)
    If LevelIndex = Mid(Form4.Combo1.Text, 1, 2) Then
        RoomNumber = Mid(ROMallHex, i * 24 + 3, 2)
        'If Val("&H" & RoomNumber) < 16 Then Form4.Command2.Enabled = True
        'If Val("&H" & RoomNumber) = 16 Then Form4.Command2.Enabled = False
        Form4.Text2.Text = "Level " & Form4.Combo1.Text & "and its room number and time data:" & vbCrLf
        Form4.Text2.Text = Form4.Text2.Text & Mid(ROMallHex, i * 24 + 1, 24) & vbCrLf
        LevelStartStream = Mid(ROMallHex, i * 24 + 1, 24)
        Form4.Text2.Text = Form4.Text2.Text & "Offset of this data: " & Hex(Val("&H" & offset_639068) + i * 12) & vbCrLf
        LevelStartStreamOffset = Hex(Val("&H" & offset_639068) + i * 12)
        Exit For
    End If
Next i
ROMallHex = ""
'**************************************************************************************************************************************************
Dim offset_78F280 As String
offset_78F280 = "78F280"
bytenum = 96

For i = LBound(ROMallbyte) + CLng(Val("&H" & offset_78F280)) To LBound(ROMallbyte) + CLng(Val("&H" & offset_78F280)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Form4.Label2.Caption = "process: Load pointers table for the Level......" & i - LBound(ROMallbyte) - CLng(Val("&H" & offset_78F280)) & "/" & bytenum
Next i

Dim BaseOffset As String
BaseOffset = Mid(ROMallHex, Val("&H" & LevelIndex) * 8 + 1 + 6, 2) & Mid(ROMallHex, Val("&H" & LevelIndex) * 8 + 1 + 4, 2) & Mid(ROMallHex, Val("&H" & LevelIndex) * 8 + 1 + 2, 2) & Mid(ROMallHex, Val("&H" & LevelIndex) * 8 + 1, 2)
BaseOffset = Hex(Val("&H" & BaseOffset) - Val("&H" & "8000000"))
LevelAllRoomPointerandDataBaseOffset = BaseOffset
Form4.Text2.Text = Form4.Text2.Text & "Base Offset for all the pointers and properties of layers for rooms：" & BaseOffset & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Offset for the pointer point to the offset：" & Hex(Val("&H" & offset_78F280) + Val("&H" & Form4.Combo1.Text) * 4) & vbCrLf
ROMallHex = ""
bytenum = 704

For i = LBound(ROMallbyte) + CLng(Val("&H" & BaseOffset)) To LBound(ROMallbyte) + CLng(Val("&H" & BaseOffset)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Form4.Label2.Caption = "process: Load pointers table for the rooms......" & i - LBound(ROMallbyte) - CLng(Val("&H" & BaseOffset)) & "/" & bytenum
Next i
LevelAllRoomPointerandDataallHex = ROMallHex

Dim tempoffset As String
For i = 0 To (Val("&H" & RoomNumber) - 1)
tempoffset = Mid(ROMallHex, 25 + 6 + i * 44 * 2, 2) & Mid(ROMallHex, 25 + 4 + i * 44 * 2, 2) & Mid(ROMallHex, 25 + 2 + i * 44 * 2, 2) & Mid(ROMallHex, 25 + i * 44 * 2, 2)
tempoffset = Val("&H" & tempoffset) - Val("&H" & "8000000")
Form4.List1.AddItem Hex(tempoffset)
Form4.List2.AddItem Hex(Val("&H" & BaseOffset) + 12 + i * 44)
tempoffset = Mid(ROMallHex, 8 + 25 + 6 + i * 44 * 2, 2) & Mid(ROMallHex, 8 + 25 + 4 + i * 44 * 2, 2) & Mid(ROMallHex, 8 + 25 + 2 + i * 44 * 2, 2) & Mid(ROMallHex, 8 + 25 + i * 44 * 2, 2)
tempoffset = Val("&H" & tempoffset) - Val("&H" & "8000000")
Form4.List3.AddItem Hex(tempoffset)
Form4.List4.AddItem Hex(Val("&H" & BaseOffset) + 12 + i * 44 + 4)
tempoffset = Mid(ROMallHex, 25 + 6 + i * 44 * 2 - 8, 2) & Mid(ROMallHex, 25 + 4 + i * 44 * 2 - 8, 2) & Mid(ROMallHex, 25 + 2 + i * 44 * 2 - 8, 2) & Mid(ROMallHex, 25 + i * 44 * 2 - 8, 2)
tempoffset = Val("&H" & tempoffset) - Val("&H" & "8000000")
Form4.List5.AddItem Hex(tempoffset)
Form4.List6.AddItem Hex(Val("&H" & BaseOffset) + 12 + i * 44 - 4)
Next i
ROMallHex = ""

Dim offset_78F21C As String           '读取每个level的Room转换信息流指针
offset_78F21C = "78F21C"
bytenum = 96

For i = LBound(ROMallbyte) + CLng(Val("&H" & offset_78F21C)) To LBound(ROMallbyte) + CLng(Val("&H" & offset_78F21C)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Form4.Label2.Caption = "process: Load data from pointer for room change info" & i - LBound(ROMallbyte) - CLng(Val("&H" & offset_78F21C)) & "/" & bytenum
Next i

Dim RoomChangeMassageOffset As String
LevelChangeRoomStreamPointerOffset = Hex(Val("&H" & offset_78F21C) + Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 4)
tempoffset = Mid(ROMallHex, Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 8 + 1 + 6, 2) & Mid(ROMallHex, Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 8 + 1 + 4, 2) & Mid(ROMallHex, Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 8 + 1 + 2, 2) & Mid(ROMallHex, Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 8 + 1, 2)
RoomChangeMassageOffset = Hex(Val("&H" & tempoffset) - Val("&H" & "8000000"))
Form4.Text2.Text = Form4.Text2.Text & "Base offset for room change info:" & RoomChangeMassageOffset & vbCrLf
LevelChangeRoomStreamOffset = RoomChangeMassageOffset
Form4.Text2.Text = Form4.Text2.Text & "the offset of the pointer:" & Hex(Val("&H" & offset_78F21C) + Val("&H" & Mid(Form4.Combo1.Text, 1, 2)) * 4) & vbCrLf
Form4TextBox2Temp = Form4.Text2.Text
ROMallHex = ""

Erase ROMallbyte()
Form4.Combo1.Enabled = True
Form4.List1.Enabled = True
Form4.List3.Enabled = True
Form4.List5.Enabled = True
End Sub

Private Sub Command1_Click()
If (Form4.Text1.Enabled = False) Or (Form4.Text1.Text = "") Then Exit Sub
Form1.Text1.Text = Form4.Text1.Text
Load Form1
Form1.Show
End Sub

Private Sub Command2_Click()   ' UNFINISHED
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

SaveDatabuffer(i + 1) = LevelAllRoomPointerandDataallHex ' & "XX101010 20000000 63223F08 63223F08 63223F08 7CDA5808 "     '缺少开始5个标志位信息，不能完成

Form11.Visible = True
End Sub

Private Sub Command3_Click()   ' UNFINISHED
If MODfilepath = "" Then
MsgBox "No MOD file Loaded", vbInformation, "Info"
Exit Sub
End If
'get last LevelRoomIndex if exist
If LevelRoomIndex = "" Then
MsgBox "You haven't chose a Room yet !", vbExclamation + vbOKOnly, "Info"
Exit Sub
End If
Dim i As Integer, j As Integer
ReDim PostlayerCompDataLength(2)
' Exclude 08601854 and 083F2263
'--------------------------------Layer 0
Form9.Text1.Text = Form9.Text1.Text & "extracting and decompressing data..." & vbCrLf
DoEvents
Dim Poffset As String
Poffset = Form4.List5.List(Val("&H" & LevelRoomIndex) - 1)
If Poffset = "601854" Then
ExistUnchangeableLayer0 = True
PostlayerCompDataLength(0) = 0
ElseIf Poffset <> "3F2263" Then
Poffset = DecompressRLE(Poffset, True)
ReDim L0_LB_000(layerWidth, layerHeight)
ReDim L0_LB_001(layerWidth, layerHeight)
PostlayerCompDataLength(0) = DataByteNumber
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
L0_LB_000(i, j) = TextMap(i, j)
Next i
Next j
Layer0Height = layerHeight
Layer0Width = layerWidth
End If
'--------------------------------Layer 1  which must be exist
Poffset = Form4.List1.List(Val("&H" & LevelRoomIndex) - 1)
If Poffset <> "3F2263" Then
Poffset = DecompressRLE(Poffset, True)
ReDim L1_LB_000(layerWidth, layerHeight)
ReDim L1_LB_001(layerWidth, layerHeight)
PostlayerCompDataLength(1) = DataByteNumber
If Form4.List5.List(Val("&H" & LevelRoomIndex) - 1) = "3F2263" Or ExistUnchangeableLayer0 = True Then
ReDim L0_LB_000(layerWidth, layerHeight)
ReDim L0_LB_001(layerWidth, layerHeight)
PostlayerCompDataLength(0) = 0
Layer0Height = layerHeight
Layer0Width = layerWidth
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
L0_LB_000(i, j) = "0000"
Next i
Next j
End If
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
L1_LB_000(i, j) = TextMap(i, j)
Next i
Next j
End If
'--------------------------------Layer 2
Poffset = Form4.List3.List(Val("&H" & LevelRoomIndex) - 1)
If Poffset <> "3F2263" Then
Poffset = DecompressRLE(Poffset, True)
ReDim L2_LB_000(layerWidth, layerHeight)
ReDim L2_LB_001(layerWidth, layerHeight)
PostlayerCompDataLength(2) = DataByteNumber
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
L2_LB_000(i, j) = TextMap(i, j)
Next i
Next j
Else
ReDim L2_LB_000(layerWidth, layerHeight)
ReDim L2_LB_001(layerWidth, layerHeight)
PostlayerCompDataLength(2) = 0
For j = 0 To layerHeight - 1
For i = 0 To layerWidth - 1
L0_LB_000(i, j) = "0000"
Next i
Next j
End If
WholeRoomChange = True
Form10.Visible = True
End Sub

Private Sub Form_Load()
Form4.Move 0, 0, 4650, 9705
Form4.Label1.FontSize = 13
Form4.Label2.FontSize = 10
Form4.Combo1.FontSize = 12

'Form11.Visible = True          'for test
'Form12.Visible = True          'for test
End Sub

Private Sub List1_Click()
If Form4.Text1.Text = "00" Then Exit Sub
Form4.List1.Enabled = False
Form4.Text1.Text = Form4.List1.Text
PointerOffset1 = Form4.List2.List(Form4.List1.ListIndex)
LevelRoomIndex = Hex(Form4.List1.ListIndex + 1)
Form4.Text2.Text = Form4TextBox2Temp & "Room Index:" & Form4.List1.ListIndex + 1 & "(Hex:" & Hex(Form4.List1.ListIndex + 1) & ")" & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Offset of pointer for Room layer 1 Data: " & PointerOffset1 & vbCrLf

Dim offset_3F2298 As String, ReadFileOffset As String
offset_3F2298 = "3F2298"

Dim FirstByte As String            '按顺序总共可以找到16个数值，一定用于各个游戏寄存器

FirstByte = Mid(LevelAllRoomPointerandDataallHex, 1 + Form4.List1.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "Tileset:" & FirstByte & vbCrLf
ReadFileOffset = Hex(Val("&H" & FirstByte) * 9 * 4 + Val("&H" & "3F2298"))

Dim TenthByte_scrollBG As String

TenthByte_scrollBG = Mid(LevelAllRoomPointerandDataallHex, 1 + 50 + Form4.List1.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "TenthByte_scrollBG register (If value = 7 then scroll):" & TenthByte_scrollBG & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping flag:" & Mid(LevelAllRoomPointerandDataallHex, 9 + Form4.List1.ListIndex * 44 * 2, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping data:" & Mid(LevelAllRoomPointerandDataallHex, 41 + 6 + Form4.List1.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 4 + Form4.List1.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 2 + Form4.List1.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + Form4.List1.ListIndex * 44 * 2, 2) & vbCrLf
'********************************************读文件过程

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
bytenum = 128
For i = LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) To LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
Next i
Erase ROMallbyte()

ROMallHex = Mid(ROMallHex, 1, 72)
Form4.Text2.Text = Form4.Text2.Text & "Tile图块所在ROM地址：" & Mid(ROMallHex, 5, 2) & Mid(ROMallHex, 3, 2) & Mid(ROMallHex, 1, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 11, 2) & Mid(ROMallHex, 9, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile调色板地址：" & Mid(ROMallHex, 21, 2) & Mid(ROMallHex, 19, 2) & Mid(ROMallHex, 17, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块地址：" & Mid(ROMallHex, 29, 2) & Mid(ROMallHex, 27, 2) & Mid(ROMallHex, 25, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 35, 2) & Mid(ROMallHex, 33, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "大Tile每个图块的属性和编号数据：" & Mid(ROMallHex, 45, 2) & Mid(ROMallHex, 43, 2) & Mid(ROMallHex, 41, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针1（好像是什么的RAW）：" & Mid(ROMallHex, 55, 2) & Mid(ROMallHex, 53, 2) & Mid(ROMallHex, 51, 2) & Mid(ROMallHex, 49, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针2：" & Mid(ROMallHex, 55 + 8, 2) & Mid(ROMallHex, 53 + 8, 2) & Mid(ROMallHex, 51 + 8, 2) & Mid(ROMallHex, 49 + 8, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "精灵调色板：" & Mid(ROMallHex, 53 + 16, 2) & Mid(ROMallHex, 51 + 16, 2) & Mid(ROMallHex, 49 + 16, 2) & vbCrLf

RoomElementFirstOffset = Hex(Val("&H" & PointerOffset1) + 16)

CameraCotrolPointerOffset = ""
CameraCotrolString = ""
RoomCameraStringPointerOffset = ""
LengthOfAllPointer = 0
    Dim OutputString As String, CheckPointer As String, j As Integer, kk As Integer, FirstPointer As String
    FirstPointer = Hex(Val("&H" & "78F540") + 4 * Val("&H" & LevelNumber))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 3))
    FirstPointer = Mid(FirstPointer, 7, 2) & Mid(FirstPointer, 5, 2) & Mid(FirstPointer, 3, 2) & Mid(FirstPointer, 1, 2)
    FirstPointer = Hex(Val("&H" & FirstPointer) - Val("&H" & "8000000"))
    CameraCotrolPointerOffset = FirstPointer
    '*********************                  pointer table pointer head is Offset_78F540
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 17 * 4 - 1))    'pretend there is 17 pointers, get all the pointers
    For i = 0 To 16
    LengthOfAllPointer = LengthOfAllPointer + 4
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For
    Next i
If Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) = "03" Then
    '*********************                  start search
    For i = 0 To 16
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For             'there is so many FF after 3F9D58 as a end flag
    CheckPointer = Mid(FirstPointer, 7 + 8 * i, 2) & Mid(FirstPointer, 5 + 8 * i, 2) & Mid(FirstPointer, 3 + 8 * i, 2) & Mid(FirstPointer, 1 + 8 * i, 2)
    CheckPointer = Hex(Val("&H" & CheckPointer) - Val("&H" & "8000000"))

    OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 1))
        If Mid(OutputString, 1, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
            RoomCameraStringPointerOffset = CheckPointer
            OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 10 * 9 + 1))
            'then go on to enumerate the camera control flag
            CameraCotrolString = Mid(OutputString, 1, 4)
            kk = Val("&H" & Mid(OutputString, 3, 2))
            If kk <> 0 Then Form4.Text2.Text = Form4.Text2.Text & "Exist Camera Control !" & vbCrLf
            For j = 0 To (kk - 1)
            CameraCotrolString = CameraCotrolString & Mid(OutputString, 18 * j + 5, 18)
            Next j
            Exit For
        End If
    Next i
End If

Form4.List1.Enabled = True
End Sub

Private Sub List1_DblClick()
Dim i As Integer
i = MsgBox("Are you sure to change this value?", vbYesNo, "Info")
If i = vbNo Then Exit Sub
Dim str As String, str1 As String
str1 = InputBox("input an Offset to ", "Info", "3F2263")
If str1 = "" Then Exit Sub
str = Replace(str1, Chr(32), "")
str = Replace(str, Chr(13), "")
str = Replace(str, Chr(10), "")
str = Right("00" & Hex(Val("&H" & "8000000") + Val("&H" & str)), 8)
str = Mid(str, 7, 2) & Mid(str, 5, 2) & Mid(str, 3, 2) & Mid(str, 1, 2)
If SaveDataOffset(97) <> "" Then
    MsgBox "buffer memory used up, save all and retry"
    Exit Sub
End If
Dim TempAddress As String
TempAddress = Form4.List2.List(Form4.List1.ListIndex)
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDataOffset(i) = TempAddress Then Exit For
Next i
SaveDataOffset(i) = TempAddress
SaveDatabuffer(i) = str
Form4.List1.AddItem str1, List1.ListIndex
Form4.List1.RemoveItem List1.ListIndex
End Sub

Private Sub List1_Scroll()
List2.TopIndex = List1.TopIndex
List3.TopIndex = List1.TopIndex
List4.TopIndex = List1.TopIndex
List5.TopIndex = List1.TopIndex
List6.TopIndex = List1.TopIndex
End Sub

Private Sub List2_Scroll()
List1.TopIndex = List2.TopIndex
List3.TopIndex = List2.TopIndex
List4.TopIndex = List2.TopIndex
List5.TopIndex = List2.TopIndex
List6.TopIndex = List2.TopIndex
End Sub

Private Sub List3_Click()
If Form4.Text1.Text = "00" Then Exit Sub
Form4.List3.Enabled = False
Form4.Text1.Text = Form4.List3.Text
PointerOffset1 = Form4.List4.List(Form4.List3.ListIndex)
LevelRoomIndex = Hex(Form4.List3.ListIndex + 1)
Form4.Text2.Text = Form4TextBox2Temp & "Room Index:" & Form4.List3.ListIndex + 1 & "(Hex:" & Hex(Form4.List3.ListIndex + 1) & ")" & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Offset of pointer for Room layer 2 Data: " & PointerOffset1 & vbCrLf

Dim offset_3F2298 As String, ReadFileOffset As String
offset_3F2298 = "3F2298"

Dim FirstByte As String

FirstByte = Mid(LevelAllRoomPointerandDataallHex, 1 + Form4.List3.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "Tileset:" & FirstByte & vbCrLf
ReadFileOffset = Hex(Val("&H" & FirstByte) * 9 * 4 + Val("&H" & "3F2298"))

Dim TenthByte_scrollBG As String

TenthByte_scrollBG = Mid(LevelAllRoomPointerandDataallHex, 1 + 50 + Form4.List3.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "TenthByte_scrollBG register (If value = 7 then scroll):" & TenthByte_scrollBG & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping flag:" & Mid(LevelAllRoomPointerandDataallHex, 9 + Form4.List3.ListIndex * 44 * 2, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping data:" & Mid(LevelAllRoomPointerandDataallHex, 41 + 6 + Form4.List3.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 4 + Form4.List3.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 2 + Form4.List3.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + Form4.List3.ListIndex * 44 * 2, 2) & vbCrLf
'********************************************读文件过程

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
bytenum = 128
For i = LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) To LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
Next i
Erase ROMallbyte()

ROMallHex = Mid(ROMallHex, 1, 72)
Form4.Text2.Text = Form4.Text2.Text & "Tile图块所在ROM地址：" & Mid(ROMallHex, 5, 2) & Mid(ROMallHex, 3, 2) & Mid(ROMallHex, 1, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 11, 2) & Mid(ROMallHex, 9, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile调色板地址：" & Mid(ROMallHex, 21, 2) & Mid(ROMallHex, 19, 2) & Mid(ROMallHex, 17, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块地址：" & Mid(ROMallHex, 29, 2) & Mid(ROMallHex, 27, 2) & Mid(ROMallHex, 25, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 35, 2) & Mid(ROMallHex, 33, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "大Tile每个图块的属性和编号数据：" & Mid(ROMallHex, 45, 2) & Mid(ROMallHex, 43, 2) & Mid(ROMallHex, 41, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针1（好像是什么的RAW）：" & Mid(ROMallHex, 55, 2) & Mid(ROMallHex, 53, 2) & Mid(ROMallHex, 51, 2) & Mid(ROMallHex, 49, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针2：" & Mid(ROMallHex, 55 + 8, 2) & Mid(ROMallHex, 53 + 8, 2) & Mid(ROMallHex, 51 + 8, 2) & Mid(ROMallHex, 49 + 8, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "精灵调色板：" & Mid(ROMallHex, 53 + 16, 2) & Mid(ROMallHex, 51 + 16, 2) & Mid(ROMallHex, 49 + 16, 2) & vbCrLf

RoomElementFirstOffset = Hex(Val("&H" & PointerOffset1) + 12)

CameraCotrolPointerOffset = ""
CameraCotrolString = ""
RoomCameraStringPointerOffset = ""
LengthOfAllPointer = 0
    Dim OutputString As String, CheckPointer As String, j As Integer, kk As Integer, FirstPointer As String
    FirstPointer = Hex(Val("&H" & "78F540") + 4 * Val("&H" & LevelNumber))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 3))
    FirstPointer = Mid(FirstPointer, 7, 2) & Mid(FirstPointer, 5, 2) & Mid(FirstPointer, 3, 2) & Mid(FirstPointer, 1, 2)
    FirstPointer = Hex(Val("&H" & FirstPointer) - Val("&H" & "8000000"))
    CameraCotrolPointerOffset = FirstPointer
    '*********************                  pointer table pointer head is Offset_78F540
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 17 * 4 - 1))    'pretend there is 17 pointers, get all the pointers
    For i = 0 To 16
    LengthOfAllPointer = LengthOfAllPointer + 4
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For
    Next i
If Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) = "03" Then
    '*********************                  start search
    For i = 0 To 16
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For             'there is so many FF after 3F9D58 as a end flag
    CheckPointer = Mid(FirstPointer, 7 + 8 * i, 2) & Mid(FirstPointer, 5 + 8 * i, 2) & Mid(FirstPointer, 3 + 8 * i, 2) & Mid(FirstPointer, 1 + 8 * i, 2)
    CheckPointer = Hex(Val("&H" & CheckPointer) - Val("&H" & "8000000"))

    OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 1))
        If Mid(OutputString, 1, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
            RoomCameraStringPointerOffset = CheckPointer
            OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 10 * 9 + 1))
            'then go on to enumerate the camera control flag
            CameraCotrolString = Mid(OutputString, 1, 4)
            kk = Val("&H" & Mid(OutputString, 3, 2))
            If kk <> 0 Then Form4.Text2.Text = Form4.Text2.Text & "Exist Camera Control !" & vbCrLf
            For j = 0 To (kk - 1)
            CameraCotrolString = CameraCotrolString & Mid(OutputString, 18 * j + 5, 18)
            Next j
            Exit For
        End If
    Next i
End If

Form4.List3.Enabled = True
End Sub

Private Sub List3_DblClick()
Dim i As Integer
i = MsgBox("Are you sure to change this value?", vbYesNo, "Info")
If i = vbNo Then Exit Sub
Dim str As String, str1 As String
str1 = InputBox("input an Offset to ", "Info", "3F2263")
If str1 = "" Then Exit Sub
str = Replace(str1, Chr(32), "")
str = Replace(str, Chr(13), "")
str = Replace(str, Chr(10), "")
str = Right("00" & Hex(Val("&H" & "8000000") + Val("&H" & str)), 8)
str = Mid(str, 7, 2) & Mid(str, 5, 2) & Mid(str, 3, 2) & Mid(str, 1, 2)
If SaveDataOffset(97) <> "" Then
    MsgBox "buffer memory used up, save all and retry"
    Exit Sub
End If
Dim TempAddress As String
TempAddress = Form4.List4.List(Form4.List3.ListIndex)
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDataOffset(i) = TempAddress Then Exit For
Next i
SaveDataOffset(i) = TempAddress
SaveDatabuffer(i) = str
Form4.List3.AddItem str1, List3.ListIndex
Form4.List3.RemoveItem List3.ListIndex
End Sub

Private Sub List3_Scroll()
List1.TopIndex = List3.TopIndex
List2.TopIndex = List3.TopIndex
List4.TopIndex = List3.TopIndex
List5.TopIndex = List3.TopIndex
List6.TopIndex = List3.TopIndex
End Sub

Private Sub List4_Scroll()
List1.TopIndex = List4.TopIndex
List2.TopIndex = List4.TopIndex
List3.TopIndex = List4.TopIndex
List5.TopIndex = List4.TopIndex
List6.TopIndex = List4.TopIndex
End Sub

Private Sub List5_Click()
If Form4.Text1.Text = "00" Then Exit Sub
Form4.List5.Enabled = False
Form4.Text1.Text = Form4.List5.Text
PointerOffset1 = Form4.List6.List(Form4.List5.ListIndex)
LevelRoomIndex = Hex(Form4.List5.ListIndex + 1)
Form4.Text2.Text = Form4TextBox2Temp & "Room Index:" & Form4.List5.ListIndex + 1 & "(Hex:" & Hex(Form4.List5.ListIndex + 1) & ")" & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Offset of pointer for Room layer 0 Data: " & PointerOffset1 & vbCrLf

Dim offset_3F2298 As String, ReadFileOffset As String
offset_3F2298 = "3F2298"

Dim FirstByte As String

FirstByte = Mid(LevelAllRoomPointerandDataallHex, 1 + Form4.List5.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "Tileset:" & FirstByte & vbCrLf
ReadFileOffset = Hex(Val("&H" & FirstByte) * 9 * 4 + Val("&H" & "3F2298"))

Dim TenthByte_scrollBG As String

TenthByte_scrollBG = Mid(LevelAllRoomPointerandDataallHex, 1 + 50 + Form4.List5.ListIndex * 44 * 2, 2)
Form4.Text2.Text = Form4.Text2.Text & "TenthByte_scrollBG register (If value = 7 then scroll):" & TenthByte_scrollBG & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping flag:" & Mid(LevelAllRoomPointerandDataallHex, 9 + Form4.List5.ListIndex * 44 * 2, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "BG pointer mapping data:" & Mid(LevelAllRoomPointerandDataallHex, 41 + 6 + Form4.List5.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 4 + Form4.List5.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + 2 + Form4.List5.ListIndex * 44 * 2, 2) & Mid(LevelAllRoomPointerandDataallHex, 41 + Form4.List5.ListIndex * 44 * 2, 2) & vbCrLf
'********************************************读文件过程

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
bytenum = 128
For i = LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) To LBound(ROMallbyte) + CLng(Val("&H" & ReadFileOffset)) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
Next i
Erase ROMallbyte()

ROMallHex = Mid(ROMallHex, 1, 72)
Form4.Text2.Text = Form4.Text2.Text & "Tile图块所在ROM地址：" & Mid(ROMallHex, 5, 2) & Mid(ROMallHex, 3, 2) & Mid(ROMallHex, 1, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 11, 2) & Mid(ROMallHex, 9, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "Tile调色板地址：" & Mid(ROMallHex, 21, 2) & Mid(ROMallHex, 19, 2) & Mid(ROMallHex, 17, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块地址：" & Mid(ROMallHex, 29, 2) & Mid(ROMallHex, 27, 2) & Mid(ROMallHex, 25, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "背景Tile图块加载长度（单位是Hex byte）：" & Mid(ROMallHex, 35, 2) & Mid(ROMallHex, 33, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "大Tile每个图块的属性和编号数据：" & Mid(ROMallHex, 45, 2) & Mid(ROMallHex, 43, 2) & Mid(ROMallHex, 41, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针1（好像是什么的RAW）：" & Mid(ROMallHex, 55, 2) & Mid(ROMallHex, 53, 2) & Mid(ROMallHex, 51, 2) & Mid(ROMallHex, 49, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "未知指针2：" & Mid(ROMallHex, 55 + 8, 2) & Mid(ROMallHex, 53 + 8, 2) & Mid(ROMallHex, 51 + 8, 2) & Mid(ROMallHex, 49 + 8, 2) & vbCrLf
Form4.Text2.Text = Form4.Text2.Text & "精灵调色板：" & Mid(ROMallHex, 53 + 16, 2) & Mid(ROMallHex, 51 + 16, 2) & Mid(ROMallHex, 49 + 16, 2) & vbCrLf

RoomElementFirstOffset = Hex(Val("&H" & PointerOffset1) + 20)

CameraCotrolPointerOffset = ""
CameraCotrolString = ""
RoomCameraStringPointerOffset = ""
LengthOfAllPointer = 0
    Dim OutputString As String, CheckPointer As String, j As Integer, kk As Integer, FirstPointer As String
    FirstPointer = Hex(Val("&H" & "78F540") + 4 * Val("&H" & LevelNumber))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 3))
    FirstPointer = Mid(FirstPointer, 7, 2) & Mid(FirstPointer, 5, 2) & Mid(FirstPointer, 3, 2) & Mid(FirstPointer, 1, 2)
    FirstPointer = Hex(Val("&H" & FirstPointer) - Val("&H" & "8000000"))
    CameraCotrolPointerOffset = FirstPointer
    '*********************                  pointer table pointer head is Offset_78F540
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 17 * 4 - 1))    'pretend there is 17 pointers, get all the pointers
    For i = 0 To 16
    LengthOfAllPointer = LengthOfAllPointer + 4
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For
    Next i
If Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) = "03" Then
    '*********************                  start search
    For i = 0 To 16
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For             'there is so many FF after 3F9D58 as a end flag
    CheckPointer = Mid(FirstPointer, 7 + 8 * i, 2) & Mid(FirstPointer, 5 + 8 * i, 2) & Mid(FirstPointer, 3 + 8 * i, 2) & Mid(FirstPointer, 1 + 8 * i, 2)
    CheckPointer = Hex(Val("&H" & CheckPointer) - Val("&H" & "8000000"))

    OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 1))
        If Mid(OutputString, 1, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
            RoomCameraStringPointerOffset = CheckPointer
            OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 10 * 9 + 1))
            'then go on to enumerate the camera control flag
            CameraCotrolString = Mid(OutputString, 1, 4)
            kk = Val("&H" & Mid(OutputString, 3, 2))
            If kk <> 0 Then Form4.Text2.Text = Form4.Text2.Text & "Exist Camera Control !" & vbCrLf
            For j = 0 To (kk - 1)
            CameraCotrolString = CameraCotrolString & Mid(OutputString, 18 * j + 5, 18)
            Next j
            Exit For
        End If
    Next i
End If

Form4.List5.Enabled = True
End Sub

Private Sub List5_DblClick()
Dim i As Integer
i = MsgBox("Are you sure to change this value?", vbYesNo, "Info")
If i = vbNo Then Exit Sub
Dim str As String, str1 As String
str1 = InputBox("input an Offset to ", "Info", "3F2263")
If str1 = "" Then Exit Sub
str = Replace(str1, Chr(32), "")
str = Replace(str, Chr(13), "")
str = Replace(str, Chr(10), "")
str = Right("00" & Hex(Val("&H" & "8000000") + Val("&H" & str)), 8)
str = Mid(str, 7, 2) & Mid(str, 5, 2) & Mid(str, 3, 2) & Mid(str, 1, 2)
If SaveDataOffset(97) <> "" Then
    MsgBox "buffer memory used up, save all and retry"
    Exit Sub
End If
Dim TempAddress As String
TempAddress = Form4.List6.List(Form4.List5.ListIndex)
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDataOffset(i) = TempAddress Then Exit For
Next i
SaveDataOffset(i) = TempAddress
SaveDatabuffer(i) = str
Form4.List5.AddItem str1, List5.ListIndex
Form4.List5.RemoveItem List5.ListIndex
End Sub

Private Sub List5_Scroll()
List1.TopIndex = List5.TopIndex
List2.TopIndex = List5.TopIndex
List3.TopIndex = List5.TopIndex
List4.TopIndex = List5.TopIndex
List6.TopIndex = List5.TopIndex
End Sub

Private Sub List6_Scroll()
List1.TopIndex = List6.TopIndex
List2.TopIndex = List6.TopIndex
List3.TopIndex = List6.TopIndex
List5.TopIndex = List6.TopIndex
List4.TopIndex = List6.TopIndex
End Sub
