VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Textmap editor"
   ClientHeight    =   10320
   ClientLeft      =   225
   ClientTop       =   675
   ClientWidth     =   18180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10320
   ScaleWidth      =   18180
   Visible         =   0   'False
   Begin VB.CommandButton Command16 
      Caption         =   "deliver to Visual Editor"
      Height          =   615
      Left            =   12840
      TabIndex        =   33
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Right"
      Height          =   375
      Left            =   13560
      TabIndex        =   32
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Left"
      Height          =   375
      Left            =   12840
      TabIndex        =   31
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "down"
      Height          =   375
      Left            =   13200
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "up"
      Height          =   375
      Left            =   13200
      TabIndex        =   29
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "fullfill bigger textmap by Flag"
      Height          =   495
      Left            =   10920
      TabIndex        =   27
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Resave three Flags"
      Height          =   735
      Left            =   8040
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "rewrite textmap in order ==>"
      Height          =   495
      Left            =   10920
      TabIndex        =   24
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   6255
      Left            =   14520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "replace one Tile  hundreds digit"
      Height          =   495
      Left            =   10920
      TabIndex        =   21
      ToolTipText     =   "直接操作layer 2 的缓存数据并对layer 1 进行一次缓存"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Save Compressed Data to File"
      Height          =   495
      Left            =   10920
      TabIndex        =   19
      Top             =   9000
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   10920
      TabIndex        =   17
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "repalce one byte"
      Height          =   495
      Left            =   10920
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "compress all"
      Height          =   495
      Left            =   10920
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   7080
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   7320
      Width           =   10695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "save HighBytes temp"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HighBytes"
      Height          =   495
      Left            =   10920
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save LowBytes temp"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10920
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LowBytes"
      Height          =   495
      Left            =   10920
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   10695
   End
   Begin VB.Label Label11 
      Caption         =   "Add Lines:"
      Height          =   375
      Left            =   12840
      TabIndex        =   28
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      Height          =   735
      Left            =   10920
      TabIndex        =   25
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "inputTextmap in special order:"
      Height          =   615
      Left            =   14520
      TabIndex        =   23
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   6960
      Width           =   9135
   End
   Begin VB.Label Label7 
      Caption         =   "Save Address："
      Height          =   255
      Left            =   11040
      TabIndex        =   18
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "start address:"
      Height          =   735
      Left            =   9360
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Flat 3[do not change]:"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Height Hex:"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Width Hex:"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1092
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_LINEFROMCHAR = &HC9

Public LineNumberinForm2Text1 As Long
Public IndexinOneLine As Long

Function LineNo(ByVal txthwnd As Long) As Long
LineNo = SendMessageLong(txthwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
LineNo = Format$(LineNo, "##.###")
End Function

Private Sub Command1_Click()
If gbafilepath = "" Then Exit Sub
Form2.Label10.Caption = "line：" & "0" & "    Number：" & "0"
Form2.Label2.Caption = "layer 1:"
Form2.Text1.Text = ""

Dim i As Long, j As Long

For j = 0 To Val("&H" & heighta2) - 1 '写列
    For i = 0 To Val("&H" & widtha1) - 1 '写行
    DoEvents
    Form2.Text1.Text = Form2.Text1.Text & Mid(Hexstream1, 1 + 2 * i + 2 * j * Val("&H" & widtha1), 2) & " "
    Next i
Form2.Text1.Text = Form2.Text1.Text & vbCrLf
Next j

Form2.Command2.Enabled = True
Form2.Command4.Enabled = False
End Sub

Private Sub Command10_Click()
widtha1 = Replace(Form2.Text3.Text, " ", "")
heighta2 = Replace(Form2.Text4.Text, " ", "")
transmita3 = Replace(Form2.Text5.Text, " ", "")
End Sub

Private Sub Command12_Click()
Dim Num1 As Integer, Hex1 As String

Num1 = InputBox("input how much line you want to add!", "info", 1)
If Num1 <= 0 Then
MsgBox "Wrong!"
Exit Sub
End If

If Val("&H" & widtha1) * (Val("&H" & heighta2) + Num1) >= Val("&H" & "FFF") Then
MsgBox "Map too large !"
Exit Sub
End If

Hex1 = Right("0000" & InputBox("input filling Word", "info", 40), 4)
Dim i As Integer, j As Integer, addstr As String
If Hexstream2 = "" Then GoTo DoOne1
For j = 1 To Num1
For i = 1 To Val("&H" & widtha1)
addstr = addstr & Mid(Hex1, 1, 2)
Next i
Next j
Hexstream2 = addstr & Hexstream2

addstr = ""
DoOne1:
For j = 1 To Num1
For i = 1 To Val("&H" & widtha1)
addstr = addstr & Mid(Hex1, 3, 2)
Next i
Next j
Hexstream1 = addstr & Hexstream1

heighta2 = Hex(Val("&H" & heighta2) + Num1)
Text4.Text = heighta2
Text1.Text = ""
MsgBox "Finish !"
End Sub

Private Sub Command13_Click()
Dim Num1 As Integer, Hex1 As String

Num1 = InputBox("input how much line you want to add!", "info", 1)
If Num1 <= 0 Then
MsgBox "wrong !"
Exit Sub
End If

If Val("&H" & widtha1) * (Val("&H" & heighta2) + Num1) >= Val("&H" & "FFF") Then
MsgBox "Map too large !"
Exit Sub
End If

Hex1 = Right("0000" & InputBox("input filling Word", "info", 40), 4)

Dim i As Integer, j As Integer, addstr As String
If Hexstream2 = "" Then GoTo DoOne2
For j = 1 To Num1
For i = 1 To Val("&H" & widtha1)
addstr = addstr & Mid(Hex1, 1, 2)
Next i
Next j
Hexstream2 = Hexstream2 & addstr

addstr = ""
DoOne2:
For j = 1 To Num1
For i = 1 To Val("&H" & widtha1)
addstr = addstr & Mid(Hex1, 3, 2)
Next i
Next j
Hexstream1 = Hexstream1 & addstr

heighta2 = Hex(Val("&H" & heighta2) + Num1)
Text4.Text = heighta2
Text1.Text = ""
MsgBox "Finish !"
End Sub

Private Sub Command14_Click()
Dim Num1 As Integer, Hex1 As String
Num1 = InputBox("input how much line you want to add!", "info", 1)
If Num1 <= 0 Then
MsgBox "Wrong!"
Exit Sub
End If

If Val("&H" & widtha1) * (Val("&H" & heighta2) + Num1) >= Val("&H" & "FFF") Then
MsgBox "Map too large !"
Exit Sub
End If

Hex1 = Right("0000" & InputBox("input filling Word", "info", 40), 4)
Dim i As Integer, j As Integer, addstr As String
If Hexstream2 = "" Then GoTo DoOne3
For i = 0 To Val("&H" & heighta2) - 1
For j = 1 To Num1
addstr = addstr & Mid(Hex1, 1, 2)
Next j
addstr = addstr & Mid(Hexstream2, 1 + 2 * Val("&H" & widtha1) * i, 2 * Val("&H" & widtha1))
Next i
Hexstream2 = addstr

addstr = ""
DoOne3:
For i = 0 To Val("&H" & heighta2) - 1
For j = 1 To Num1
addstr = addstr & Mid(Hex1, 1, 2)
Next j
addstr = addstr & Mid(Hexstream1, 1 + 2 * Val("&H" & widtha1) * i, 2 * Val("&H" & widtha1))
Next i
Hexstream1 = addstr

widtha1 = Hex(Val("&H" & widtha1) + Num1)
Text3.Text = widtha1
Text1.Text = ""
MsgBox "Finish !"
End Sub

Private Sub Command15_Click()
Dim Num1 As Integer, Hex1 As String
Num1 = InputBox("input how much line you want to add!", "info", 1)
If Num1 <= 0 Then
MsgBox "Wrong!"
Exit Sub
End If

If Val("&H" & widtha1) * (Val("&H" & heighta2) + Num1) >= Val("&H" & "FFF") Then
MsgBox "Map too large !"
Exit Sub
End If

Hex1 = Right("0000" & InputBox("input filling Word", "info", 40), 4)
Dim i As Integer, j As Integer, addstr As String
If Hexstream2 = "" Then GoTo DoOne3
For i = 0 To Val("&H" & heighta2) - 1
addstr = addstr & Mid(Hexstream2, 1 + 2 * Val("&H" & widtha1) * i, 2 * Val("&H" & widtha1))
For j = 1 To Num1
addstr = addstr & Mid(Hex1, 1, 2)
Next j
Next i
Hexstream2 = addstr

addstr = ""
DoOne3:
For i = 0 To Val("&H" & heighta2) - 1
addstr = addstr & Mid(Hexstream1, 1 + 2 * Val("&H" & widtha1) * i, 2 * Val("&H" & widtha1))
For j = 1 To Num1
addstr = addstr & Mid(Hex1, 1, 2)
Next j
Next i
Hexstream1 = addstr

widtha1 = Hex(Val("&H" & widtha1) + Num1)
Text3.Text = widtha1
Text1.Text = ""
MsgBox "Finish !"
End Sub

Private Sub Command16_Click()
If MODfilepath = "" Then
MsgBox "No MOD file Loaded", vbInformation, "Info"
Exit Sub
End If
IsDeliver = True
Form10.Visible = True
End Sub

Private Sub Command2_Click()
Dim strtext As String

strtext = Form2.Text1.Text

strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Hexstream1 = strtext
End Sub

Private Sub Command3_Click()
If gbafilepath = "" Then Exit Sub
Form2.Label10.Caption = "line：" & "0" & "    Number：" & "0"
Form2.Label2.Caption = "layer 2:"
Form2.Text1.Text = ""

Dim i As Long, j As Long

For j = 0 To Val("&H" & heighta2) - 1 '写列
    For i = 0 To Val("&H" & widtha1) - 1 '写行
    DoEvents
    Form2.Text1.Text = Form2.Text1.Text & Mid(Hexstream2, 1 + 2 * i + 2 * j * Val("&H" & widtha1), 2) & " "
    Next i
Form2.Text1.Text = Form2.Text1.Text & vbCrLf
Next j

Form2.Command4.Enabled = True
Form2.Command2.Enabled = False
End Sub

Private Sub Command4_Click()
Dim strtext As String

strtext = Form2.Text1.Text

strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Hexstream2 = strtext
End Sub

Private Sub Command5_Click()
If gbafilepath = "" Then Exit Sub

Form2.Text6.Text = ""
widtha1 = Replace(Form2.Text3.Text, " ", "")
heighta2 = Replace(Form2.Text4.Text, " ", "")
transmita3 = Replace(Form2.Text5.Text, " ", "")

Dim ALLCOMPRESSDATA As String

ALLCOMPRESSDATA = ALLCOMPRESSDATA & Right("00" & widtha1, 2) & Right("00" & heighta2, 2) & Right("00" & transmita3, 2)
ALLCOMPRESSDATA = ALLCOMPRESSDATA & CompressDataOnly(Hexstream1)

'If Len(ALLCOMPRESSDATA) > layer1compressdatalength - 8 Then
'MsgBox "layer 1 压缩数据长度超出原压缩数据，若要写入请重新修改ROM的指针。", vbOKOnly, "Warning"
'End If
Dim transleftzero1 As Long
Dim i As Long, j As Long, no1 As Boolean
no1 = True
For j = 0 To CLng("&H" & heighta2) - 1
For i = 0 To CLng("&H" & widtha1) - 1
If Mid(Hexstream2, 2 * i + 2 * j * Val("&H" & widtha1) + 1, 2) <> "00" Then
no1 = False
End If
Next i
Next j
If no1 = True Then
Hexstream2 = ""
End If

If Hexstream2 = "" Then
If leftzerozero1 > 6 Then
transleftzero1 = 6
Else
transleftzero1 = leftzerozero1
End If
    
    If Len(ALLCOMPRESSDATA) > layer1compressdatalength + 6 - 2 * transleftzero1 Then '尽量别把关卡要素和怪物放在Room边缘
    MsgBox "New Data too long, new offset should be input, you can do it yourself or the porgram will do it automatically！", vbOKOnly, "Warning"
        If layer1compressdatalength = 0 Then MsgBox "making new Layer needs you to change pointer in yourself.", vbOKOnly + vbInformation, "information!"
    Else
    Form2.Text6.Text = startoffset
    End If

'to get rid of the difficulty in meet with the lowest digit and second lowest digit are all equal to 0
If (Val("&H" & "8000") + Val("&H" & widtha1) * Val("&H" & heighta2) Mod 256) = 0 Then
ALLCOMPRESSDATA = ALLCOMPRESSDATA & "0002" & Right(Hex(CLng("&H" & "8000") + CLng("&H" & widtha1) * CLng("&H" & heighta2) + 10), 4)
Else
ALLCOMPRESSDATA = ALLCOMPRESSDATA & "0002" & Right(Hex(CLng("&H" & "8000") + CLng("&H" & widtha1) * CLng("&H" & heighta2)), 4)
End If

ALLCOMPRESSDATA = ALLCOMPRESSDATA & "000000"
Do
    If Len(ALLCOMPRESSDATA) < layer1compressdatalength + 6 + 8 Then
        ALLCOMPRESSDATA = ALLCOMPRESSDATA & "00"
    Else
    Exit Do
    End If
Loop

Form2.Text2.Text = ALLCOMPRESSDATA
Form2.Label8.Caption = "bytes：" & Len(Form2.Text2.Text) / 2
Exit Sub
End If

ALLCOMPRESSDATA = ALLCOMPRESSDATA & "0001"
ALLCOMPRESSDATA = ALLCOMPRESSDATA & CompressDataOnly(Hexstream2)
'**********************************************缺少后续判断数据部分
If Len(ALLCOMPRESSDATA) + 6 > layer2compressdatalength + layer1compressdatalength + 6 + 4 Then
MsgBox "new offset for inport, you can do it yourself or the porgram will do it automatically！", vbOKOnly, "Warning"
Else
Form2.Text6.Text = startoffset
End If

Do
    If Len(ALLCOMPRESSDATA) < layer1compressdatalength + layer2compressdatalength + 4 Then
        ALLCOMPRESSDATA = ALLCOMPRESSDATA & "00"
    Else
    Exit Do
    End If
Loop

ALLCOMPRESSDATA = ALLCOMPRESSDATA & "000000"

Form2.Text2.Text = ALLCOMPRESSDATA
Form2.Label8.Caption = "Bytes：" & Len(Form2.Text2.Text) / 2
End Sub

Private Sub Command6_Click()
If Form2.Command4.Enabled = True Then Exit Sub
If Hexstream2 = "" Then
MsgBox " layer 2 does not exist!", vbInformation, "info"
Exit Sub
End If
Dim strtext As String
strtext = Form2.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")
Hexstream1 = strtext

Dim ChangeByte As String, i As Long, ChangeToHundredsDigit As String
ChangeByte = Right("00" & InputBox("输入要修改百位的layer 1 字节", "提示", 0), 2)
ChangeToHundredsDigit = Right("00" & InputBox("要把百位替换成多少？", "提示", 0), 2)
For i = 1 To (Len(Hexstream1) - 1) Step 2
If Mid(Hexstream1, i, 2) = ChangeByte Then Mid(Hexstream2, i, 2) = ChangeToHundredsDigit
Next i

MsgBox "finish change！"
End Sub

Private Sub Command7_Click()
Dim inputstr1 As String, inputstr2 As String
inputstr1 = Right("00" & CStr(InputBox("input what you want to change in the textbox", "inform", 0)), 2)
inputstr2 = Right("00" & CStr(InputBox("input what to change to", "inform", 0)), 2)

For i = 1 To Len(Hexstream1) / 2 + 1
Text1.Text = Replace(Text1.Text, inputstr1, inputstr2)
Next
End Sub

Private Sub Command8_Click()
If gbafilepath = "" Then Exit Sub
Dim i As Integer, j As Long
Dim str1 As String
Dim maxnum As Long
Dim TempAddress As Long
If IfisNewRoom = True Then
i = MsgBox("make sure you are making new Room！！！", vbYesNo, "info")
If i <> vbYes Then Exit Sub
    If SaveDataOffset(98) <> "" Then
        MsgBox "buffer memory used up, save all and retry !"
        Exit Sub
    End If
    For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
    Next i
SaveDataOffset(i) = SaveDatabuffer(0)
SaveDatabuffer(i) = Form2.Text2.Text & "FF"
'MsgBox "保存后请自行修改 78F970 处的下次写入地址！"
Form2.Label8.Caption = "Automatically make the offset"
Form2.Text6.Text = SaveDatabuffer(0)
SaveDataOffset(i + 1) = PointerOffset1
TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDatabuffer(0))
SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
TempAddress = Val("&H" & SaveDatabuffer(0)) + Len(Form2.Text2.Text) / 2 + 1
SaveDatabuffer(0) = Right("00" & Hex(TempAddress), 8)

Exit Sub
End If

If Form2.Text6.Text = "" Then         '检查Textbox，就是说可以自己输入地址，但是注意要小于顺序写入地址，即第一条写入记录
    If SaveDataOffset(98) <> "" Then
        MsgBox "buffer memory used up, save all and retry !"
        Exit Sub
    End If
    TempAddress = Val("&H" & SaveDatabuffer(0)) + Len(Form2.Text2.Text) / 2
    For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
        If SaveDataOffset(i) = Right("00" & Hex(TempAddress), 8) Then
        MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
        Exit Sub
        End If
    Next i
    Dim returnstr As String
    Form2.Label8.Caption = "正在搜索原储存压缩数据的地址段有无足够空位。。。"
    returnstr = FindSpace(gbafilepath, "598EEC", "59F291", "00", Len(Form2.Text2.Text) / 2 + 6)
    If returnstr = "FFFFFFFF" Then
    Form2.Label8.Caption = "正在搜索由程序写入过的地址段有无删除数据后产生的足够空位。。。"
    returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(Form2.Text2.Text) / 2)
    End If
    
    '-----------------------------------------------------------------------------------------------------------------出问题的地址在此处列出
    If Val("&H" & returnstr) >= Val("&H" & "59AD20") And Val("&H" & returnstr) <= Val("&H" & "59AE63") Then
    returnstr = FindSpace(gbafilepath, "59AE63", "59F291", "00", Len(Form2.Text2.Text) / 2 + 6)
    If returnstr = "FFFFFFFF" Then
    Form2.Label8.Caption = "正在搜索由程序写入过的地址段有无删除数据后产生的足够空位。。。"
    returnstr = FindSpace(gbafilepath, "78F97F", SaveDatabuffer(0), "00", 6 + Len(Form2.Text2.Text) / 2)
    End If
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    If returnstr = "FFFFFFFF" Then
        GoTo ReCreatNewOffset
    Else
        SaveDataOffset(i) = Hex(Val("&H" + returnstr))
        SaveDatabuffer(i) = "000000" & Form2.Text2.Text
        Form2.Text6.Text = Hex(Val("&H" + returnstr))
        SaveDataOffset(i + 1) = PointerOffset1
        TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
        SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
        SaveDataOffset(i + 2) = startoffset

        If Hexstream2 = "" Then
            maxnum = layer1compressdatalength / 2 + 7        'layer1compressdatalength / 2 + 3 + 2 + 2
        Else
            maxnum = layer1compressdatalength / 2 + layer2compressdatalength / 2 + 5    'layer1compressdatalength / 2 + layer2compressdatalength / 2 + 3 + 2
        End If
        str1 = ""
        For j = 1 To maxnum      '生成指定长度填充字节00
            str1 = str1 + "00"
        Next j
        SaveDatabuffer(i + 2) = str1
        GoTo GoldFingerFinding
    End If
ReCreatNewOffset:
    Form2.Label8.Caption = "没有找到足够空位，生成新地址"
    SaveDataOffset(i) = SaveDatabuffer(0)
    SaveDatabuffer(i) = "000000" & Form2.Text2.Text
    Form2.Text6.Text = SaveDataOffset(i)
    SaveDatabuffer(0) = Right("00" & Hex(TempAddress + 3), 8)
    SaveDataOffset(i + 1) = PointerOffset1
    TempAddress = Val("&H" & "8000000") + Val("&H" & SaveDataOffset(i)) + 3
    SaveDatabuffer(i + 1) = Mid(Right("00" & Hex(TempAddress), 8), 7, 2) & Mid(Right("00" & Hex(TempAddress), 8), 5, 2) & Mid(Right("00" & Hex(TempAddress), 8), 3, 2) & Mid(Right("00" & Hex(TempAddress), 8), 1, 2)
    SaveDataOffset(i + 2) = startoffset
    If Hexstream2 = "" Then
    maxnum = layer1compressdatalength / 2 + 7        'layer1compressdatalength / 2 + 3 + 2 + 2
    Else
    maxnum = layer1compressdatalength / 2 + layer2compressdatalength / 2 + 5    'layer1compressdatalength / 2 + layer2compressdatalength / 2 + 3 + 2
    End If
    str1 = ""
    For j = 1 To maxnum      '生成指定长度填充字节00
    str1 = str1 + "00"
    Next j
    SaveDatabuffer(i + 2) = str1
Else
    If SaveDataOffset(100) <> "" Then   '如果i=100的时候才有Null，那么记录数不够，退出该过程
        MsgBox "buffer memory used up, save all and retry !"
        Hexstream1 = ""
        Hexstream2 = ""
        widtha1 = ""
        heighta2 = ""
        transmita3 = ""
        leftzerozero1 = 0
        layer1compressdatalength = 0
        layer2compressdatalength = 0
        startoffset = ""
        PointerOffset1 = ""
        Exit Sub
    End If
    For i = 1 To 100
        If SaveDataOffset(i) = "" And i < 100 Then Exit For
        If SaveDataOffset(i) = Right("00" & Form2.Text6.Text, 8) Then
        MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
        Hexstream1 = ""
        Hexstream2 = ""
        widtha1 = ""
        heighta2 = ""
        transmita3 = ""
        leftzerozero1 = 0
        layer1compressdatalength = 0
        layer2compressdatalength = 0
        startoffset = ""
        PointerOffset1 = ""
        Exit Sub
        End If
    Next i
    SaveDataOffset(i) = startoffset
    SaveDatabuffer(i) = Form2.Text2.Text
End If
Hexstream1 = ""
Hexstream2 = ""
widtha1 = ""
heighta2 = ""
transmita3 = ""
leftzerozero1 = 0
layer1compressdatalength = 0
layer2compressdatalength = 0
startoffset = ""
PointerOffset1 = ""


GoldFingerFinding:
'添加ROOM修改记录，输出ROOM转换点用于查看
Form9.Text1.Text = Form9.Text1.Text & "new save room Room(jump to room with gold finger with offset 03000025)：" & LevelRoomIndex & vbCrLf
If IfisNewRoomConnectionDataBuffer = faise Then
Dim MessageStream As String
Dim checkStream As String
MessageStream = ReadFileHex(gbafilepath, LevelChangeRoomStreamOffset, Right("0000" & Hex(Val("&H" & LevelChangeRoomStreamOffset) + 1024), 8))
For i = 0 To 50
checkStream = checkStream & Mid(MessageStream, i * 24 + 1, 24)
If Mid(MessageStream, i * 24 + 1, 24) = "000000000000000000000000" Then Exit For
Next i
RoomConnectionDataBuffer = checkStream
End If
For j = 0 To i
If Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) = Mid(checkStream, j * 24 + 3, 2) Then
Form9.Text1.Text = Form9.Text1.Text & Right("00" & Hex(j), 2) & "   " & Mid(checkStream, j * 24 + 1, 24) & vbCrLf
End If
Next j
End Sub

Private Sub Command9_Click()
Dim i As Long, j As Integer, TextPerLength() As String
ReDim TextPerLength(700)
TextPerLength = Split(Form2.Text7.Text, vbCrLf)

'following sub is for debuging
'Form2.Text7.Text = ""
'For i = LBound(TextPerLength()) To UBound(TextPerLength())
'If TextPerLength(i) = "" And i <> 0 Then Exit For
'Form2.Text7.Text = Form2.Text7.Text & TextPerLength(i) & vbCrLf
'Next i

'this sub replace " " with ""
For i = LBound(TextPerLength()) To UBound(TextPerLength())
If TextPerLength(i) = "" And i <> 0 Then Exit For
TextPerLength(i) = Replace(TextPerLength(i), " ", "")
Next i

Dim strtext As String
strtext = Form2.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Form2.Text1.Text = ""
'input to whole textmap
j = Val("&H" & widtha1) * 2 * LineNumberinForm2Text1  '字符串中取完行数
j = j + 2 * IndexinOneLine    '取到列位置

For i = LBound(TextPerLength()) To UBound(TextPerLength())
If TextPerLength(i) = "" And i <> 0 Then Exit For
Mid(strtext, j + 1, Len(TextPerLength(i))) = TextPerLength(i)
If i = UBound(TextPerLength()) Then Exit For
j = j + Val("&H" & widtha1) * 2
Next i

Erase TextPerLength()


For j = 0 To Val("&H" & heighta2) - 1 '写列
    For i = 0 To Val("&H" & widtha1) - 1 '写行
    DoEvents
    Form2.Text1.Text = Form2.Text1.Text & Mid(strtext, 1 + 2 * i + 2 * j * Val("&H" & widtha1), 2) & " "
    Next i
Form2.Text1.Text = Form2.Text1.Text & vbCrLf
Next j
End Sub

Private Sub Form_Activate()
Form2.Text6.Text = ""

If IfisNewRoom = True Then
    Form2.Caption = "new Layer"
    Form2.Label6.Caption = "Start address: "
    Form2.Text3.Text = ""
    Form2.Text4.Text = ""
    Form2.Text5.Text = "01"

    widtha1 = ""
    heighta2 = ""

    Hexstream1 = ""
    Hexstream2 = ""

    layer1compressdatalength = 0
    layer2compressdatalength = 0
    startoffset = ""

    GoTo resizefrm2
End If

Form2.Caption = "Textmap editor"
Form2.Label6.Caption = "Start address:" & startoffset
Form2.Text3.Text = widtha1
Form2.Text4.Text = heighta2
Form2.Text5.Text = transmita3
If transmita3 = "02" Then MsgBox "only support comptress in mode 01(flag3=01)，please change Flag3 to 01"

If Hexstream1 = "" Then Text5.Text = "01"

resizefrm2:
Form2.Move 4650, 0, 18420, 10905

Form2.Text3.FontSize = 15
Form2.Text4.FontSize = 15
Form2.Text5.FontSize = 15
Form2.Text1.FontSize = 12
Form2.Text2.FontSize = 10
Form2.Text6.FontSize = 12
End Sub

Private Sub Text1_Click()
If (Form2.Command4.Enabled = True) And (Len(Hexstream2) = 0) Then Exit Sub
LineNumberinForm2Text1 = LineNo(Form2.Text1.hwnd) - 1
IndexinOneLine = (Form2.Text1.SelStart - (Val("&H" & widtha1) * 3 + 2) * LineNumberinForm2Text1) \ 3
Form2.Label10.Caption = "line：" & LineNumberinForm2Text1 & "    Number：" & IndexinOneLine
End Sub

Private Sub Text1_Keypress(KeyAscii As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 And KeyCode <> 32 And KeyCode <> 13 And KeyCode <> 10 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
KeyCode = 0 'All other letters are unwanted.
End If
End Sub

Private Sub Text3_Change()
widtha1 = Form2.Text3.Text
If Val("&H" & widtha1) * Val("&H" & heighta2) >= Val("&H" & "FFF") Then MsgBox "too large, please change", vbOKOnly + vbExclamation
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
KeyCode = 0 'All other letters are unwanted.
End If
End Sub

Private Sub Text4_Change()
heighta2 = Form2.Text4.Text
If Val("&H" & widtha1) * Val("&H" & heighta2) >= Val("&H" & "FFF") Then MsgBox "too large, please change", vbOKOnly + vbExclamation
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
KeyCode = 0 'All other letters are unwanted.
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
KeyCode = 0 'All other letters are unwanted.
End If
End Sub
