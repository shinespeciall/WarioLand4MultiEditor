VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Room property and flag"
   ClientHeight    =   8325
   ClientLeft      =   6030
   ClientTop       =   3180
   ClientWidth     =   5310
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   5310
   Visible         =   0   'False
   Begin VB.CommandButton Command10 
      Caption         =   "save without cover old data"
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "save"
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "save"
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   390
      Left            =   2880
      TabIndex        =   21
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "save"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "save"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "save"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "save"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form6.frx":0000
      Left            =   2880
      List            =   "Form6.frx":000D
      TabIndex        =   0
      Text            =   "Hard Mode"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label8 
      Caption         =   "Tileset Index:"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Layer 3 Visible:00=invisible 10=Visible"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Layer 2 Visible:00=invisible 10=Visible"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Layer 1 Visible:00=invisible 10=Visible 22=Scroll"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Layer 3 mode flag low byte: highest bit: transparent || low two bit: priority"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Camera Control: 03=exist 01=change height when wario go out of the camera 02=No control in both demention"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Scroll BG:01=NO, 07=Yes  03=no exist"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Form6TextBox1Temp As String

Public CameraCotrolString As String
Public CameraCotrolPointerOffset As String      '存放（指向指针表表头位置的指针）的地址
Public RoomCameraStringPointerOffset As String     '存放（指向Room的Camera控制流字符串的指针）的地址
Public LengthOfAllPointer As Long               '指针表总长，单位是Byte


Private Sub Combo1_Click()
Form6.Text1.Text = ""
If gbafilepath = "" Then Exit Sub
Dim TempString As String
TempString = ReadFileHex(gbafilepath, Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex), Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex + 3))  '我假设可以加载32个元素
TempString = Mid(TempString, 7, 2) & Mid(TempString, 5, 2) & Mid(TempString, 3, 2) & Mid(TempString, 1, 2)
TempString = Hex(Val("&H" & TempString) - Val("&H" & "8000000"))
RoomElementOffset = TempString
TempString = ReadFileHex(gbafilepath, TempString, Hex(Val("&H" & TempString) + Val("&H" & "20") * 4))
Dim i As Integer
i = 0
Do
Form6.Text1.Text = Form6.Text1.Text & Mid(TempString, 6 * i + 1, 6) & vbCrLf
If Mid(TempString, 6 * i + 1, 6) = "FFFFFF" Then Exit Do
i = i + 1
Loop

Form6TextBox1Temp = Replace(Form6.Text1.Text, Chr(32), "")
Form6TextBox1Temp = Replace(Form6TextBox1Temp, Chr(13), "")
Form6TextBox1Temp = Replace(Form6TextBox1Temp, Chr(10), "")
End Sub

Private Sub Command1_Click()
If gbafilepath = "" Then Exit Sub
Dim strtext As String
strtext = Form6.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Dim i As Integer, j As Long
Dim str1 As String
Dim maxnum As Long
If Form6.Text1.Text = "" Then Exit Sub        '检查Textbox，就是说可以自己输入地址，但是注意要小于顺序写入地址，即第一条写入记录
If SaveDataOffset(98) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim TempAddress As Long
TempAddress = Val("&H" & SaveDatabuffer(0)) + Len(strtext) / 2
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDatabuffer(i) = "000000" & strtext Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
    End If
Next i

If Len(strtext) <= Len(Form6TextBox1Temp) Then
    SaveDataOffset(i) = RoomElementOffset
    str1 = ""
        For j = 1 To Len(Form6TextBox1Temp) / 6    '生成指定长度填充字节00
        str1 = str1 + "000000"
        Next j
    SaveDatabuffer(i) = str1
    SaveDataOffset(i + 1) = RoomElementOffset
    SaveDatabuffer(i + 1) = strtext
    Exit Sub
End If
    Dim returnstr As String
    returnstr = FindSpace(gbafilepath, "598EEC", "59F291", "00", Len(strtext) / 2 + 12)
    If returnstr = "FFFFFFFF" Then
    returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(strtext) / 2 + 12)
    End If
    If returnstr = "FFFFFFFF" Then
        GoTo ReCreatNewOffset
    Else
        SaveDataOffset(i) = Hex(Val("&H" + returnstr))
        SaveDatabuffer(i) = "000000" & strtext
        SaveDataOffset(i + 1) = Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex)
        TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
        SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
        SaveDataOffset(i + 2) = RoomElementOffset
        str1 = ""
        For j = 1 To Len(Form6TextBox1Temp) / 6    '生成指定长度填充字节00
            str1 = str1 + "000000"
        Next j
        If Form6TextBox1Temp = "FFFFFF" Then str1 = "000000"
        SaveDatabuffer(i + 2) = str1
        Exit Sub
    End If
ReCreatNewOffset:
    SaveDataOffset(i) = SaveDatabuffer(0)
    SaveDatabuffer(i) = "000000" & strtext
    SaveDatabuffer(0) = Right("00" & Hex(TempAddress + 3), 8)
    SaveDataOffset(i + 1) = Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex)
    TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
    SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    SaveDataOffset(i + 2) = RoomElementOffset
    
    str1 = ""
    For j = 1 To Len(Form6TextBox1Temp) / 6      '生成指定长度填充字节00
    str1 = str1 + "000000"
    Next j
    If Form6TextBox1Temp = "FFFFFF" Then str1 = "000000"
    SaveDatabuffer(i + 2) = str1
End Sub

Private Sub Command10_Click()
If gbafilepath = "" Then Exit Sub
Dim strtext As String
strtext = Form6.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Dim i As Integer, j As Long
Dim str1 As String
Dim maxnum As Long
If Form6.Text1.Text = "" Then Exit Sub        '检查Textbox，就是说可以自己输入地址，但是注意要小于顺序写入地址，即第一条写入记录
If SaveDataOffset(98) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim TempAddress As Long
TempAddress = Val("&H" & SaveDatabuffer(0)) + Len(strtext) / 2
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDatabuffer(i) = "000000" & strtext Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
    End If
Next i

    Dim returnstr As String
    returnstr = FindSpace(gbafilepath, "598EEC", "59F291", "00", Len(strtext) / 2 + 12)
    If returnstr = "FFFFFFFF" Then
    returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(strtext) / 2 + 12)
    End If
    If returnstr = "FFFFFFFF" Then
        GoTo ReCreatNewOffset001
    Else
        SaveDataOffset(i) = Hex(Val("&H" + returnstr))
        SaveDatabuffer(i) = "000000" & strtext
        SaveDataOffset(i + 1) = Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex)
        TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
        SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
        Exit Sub
    End If
ReCreatNewOffset001:
    SaveDataOffset(i) = SaveDatabuffer(0)
    SaveDatabuffer(i) = "000000" & strtext
    SaveDatabuffer(0) = Right("00" & Hex(TempAddress + 3), 8)
    SaveDataOffset(i + 1) = Hex(Val("&H" & RoomElementFirstOffset) + 4 * Combo1.ListIndex)
    TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
    SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
End Sub

Private Sub Command2_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text2.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 25 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text2.Text
End Sub

Private Sub Command3_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text2.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(95) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 24 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i
SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text3.Text
If Form6.Text3.Text <> "03" Then Exit Sub

i = i + 1
Dim StrTemp As String, TempPointer As String

StrTemp = Replace(Form6.Text4.Text, Chr(32), "")
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
Form9.Text1.Text = Form9.Text1.Text & "Save Temp successfully!!" & vbCrLf
End Sub

Private Sub Command5_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text2.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 26 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text5.Text
End Sub

Private Sub Command6_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text6.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 2 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text6.Text
End Sub

Private Sub Command7_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text7.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 3 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text7.Text
End Sub

Private Sub Command8_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text8.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + 1 + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text8.Text
End Sub

Private Sub Command9_Click()
If gbafilepath = "" Then Exit Sub
If Len(Form6.Text9.Text) <> 2 Then
MsgBox "Wrong Const!"
Exit Sub
End If
If SaveDataOffset(100) <> "" Then
    MsgBox "记录条数不够，请保存所有修改记录后再使用缓存！"
    Exit Sub
End If
Dim i As Integer, TempAddress As Long
TempAddress = Val("&H" & LevelAllRoomPointerandDataBaseOffset) + (Val("&H" & LevelRoomIndex) - 1) * 44
For i = 1 To 100
If SaveDataOffset(i) = "" Then Exit For
If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
End If
Next i

SaveDataOffset(i) = Hex(TempAddress)
SaveDatabuffer(i) = Form6.Text9.Text
End Sub

Private Sub Form_Activate()
Form6.Move 4650, 1000, 5550, 8910
If LevelRoomIndex = "" Then Exit Sub
Form6.Text4.Text = ""
If LevelAllRoomPointerandDataallHex = "" Then
Form6.Visible = False
Exit Sub
End If
Form6.Label1.Caption = "Level Room Index:" & LevelRoomIndex
Form6.Text9.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)        'First byte flag
Form6.Text8.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 2 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)      'Second byte flag    Layer 3 Visible Flag
Form6.Text6.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 4 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)      'Third byte flag    Layer 1 Visible Flag
Form6.Text7.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 6 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)      'Fourth byte flag    Layer 2 Visible Flag
Form6.Text2.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 50 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)      'tenth byte flag
Form6.Text3.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)      'nineth byte flag
Form6.Text5.Text = Mid(LevelAllRoomPointerandDataallHex, 1 + 52 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2)     'eleventh byte flag

    Dim FirstPointer As String
    FirstPointer = Hex(Val("&H" & "78F540") + 4 * Val("&H" & LevelNumber))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 3))
    FirstPointer = Mid(FirstPointer, 7, 2) & Mid(FirstPointer, 5, 2) & Mid(FirstPointer, 3, 2) & Mid(FirstPointer, 1, 2)
    FirstPointer = Hex(Val("&H" & FirstPointer) - Val("&H" & "8000000"))
    CameraCotrolPointerOffset = FirstPointer
    
'If Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) = "03" Then
    '*********************                  pointer table pointer head is Offset_78F540
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 17 * 4 - 1))    'pretend there is 17 pointers, get all the pointers
    '*********************                  开始搜索
    Dim i As Integer, OutputString As String, CheckPointer As String, j As Integer, kk As Integer
    For i = 0 To 16
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For             'there is so many FF after 3F9D58 as a end flag
    CheckPointer = Mid(FirstPointer, 7 + 8 * i, 2) & Mid(FirstPointer, 5 + 8 * i, 2) & Mid(FirstPointer, 3 + 8 * i, 2) & Mid(FirstPointer, 1 + 8 * i, 2)
    CheckPointer = Hex(Val("&H" & CheckPointer) - Val("&H" & "8000000"))
    
    OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 1))
        If Mid(OutputString, 1, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
            RoomCameraStringPointerOffset = CheckPointer
            OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 10 * 9 + 1))
            'then go on to enumerate the camera control flag
            Form6.Text4.Text = Mid(OutputString, 1, 4) & vbCrLf
            CameraCotrolString = Mid(OutputString, 1, 4)
            kk = Val("&H" & Mid(OutputString, 3, 2))
            For j = 0 To (kk - 1)
            Form6.Text4.Text = Form6.Text4.Text & Mid(OutputString, 18 * j + 5, 10) & vbCrLf
            Form6.Text4.Text = Form6.Text4.Text & Mid(OutputString, 18 * j + 15, 8) & vbCrLf
            CameraCotrolString = CameraCotrolString & Mid(OutputString, 18 * j + 5, 18)
            Next j
            Exit For
        End If
    Next i
    
    LengthOfAllPointer = 0
    For i = 0 To 16
    LengthOfAllPointer = LengthOfAllPointer + 4
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For
    Next i
'End If
End Sub
