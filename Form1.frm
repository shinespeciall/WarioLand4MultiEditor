VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MAP decompress"
   ClientHeight    =   10950
   ClientLeft      =   1110
   ClientTop       =   1215
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   18270
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "clear textbox"
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear board"
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   9405
      Left            =   0
      ScaleHeight     =   9345
      ScaleWidth      =   18075
      TabIndex        =   3
      Top             =   1440
      Width           =   18135
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   17655
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9375
         Left            =   17640
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   12000
         Left            =   0
         ScaleHeight     =   11940
         ScaleWidth      =   99945
         TabIndex        =   6
         Top             =   360
         Width           =   1.00000e5
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "decompress"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "information"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   14415
   End
   Begin VB.Label Label1 
      Caption         =   "compressed map data offset"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo errorhandle

If Form1.Text1.Text = "601854" Then
MsgBox "This Layer cannot be decompressed!", vbInformation, "Info"
Exit Sub
End If

startoffset = Form1.Text1.Text
Form1.Text1.Enabled = False

Hexstream1 = ""       ' clear Room decompressed all Hex stream
Hexstream2 = ""
layer1compressdatalength = 0
layer1compressdatalength = 0


Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim nowoffset As Long    '记录偏移地址
Dim ROMallHex As String

Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, CLng(Val("&H" & Form1.Text1.Text)) + 1, ROMallbyte  'ROMallstr now contains all of the text in the file
Close #1

Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
nowoffset = 0
bytenum = 5120       'bytenum = 2048 + 2048 + 1024

For i = LBound(ROMallbyte) To LBound(ROMallbyte) + CLng(bytenum)
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
DoEvents
Form1.Label2.Caption = "information: Opening and reading" & i - LBound(ROMallbyte) & "/" & bytenum
Next i

Form1.Label2.Caption = "infprmation: decompressing..."
Dim a1 As String, a2 As String, a3 As String
Dim totlesize As Long

return1:

DoEvents
a1 = Mid(ROMallHex, Val(nowoffset) + 1, 2)
DoEvents
a2 = Mid(ROMallHex, Val(nowoffset) + 3, 2)
DoEvents
a3 = Mid(ROMallHex, Val(nowoffset) + 5, 2)
totlesize = Val("&H" & a1) * Val("&H" & a2)
Form1.Picture1.Print "start offset of compressed data:" & startoffset & "                 room 初始参数：" & a1 & a2 & a3
Form1.Picture1.Print "room width Hex:" & a1 & "    room height Hex:" & a2 & "      Decompress Type:   " & a3

nowoffset = nowoffset + 6

'确保 a1 为room宽，a2 为room高
widtha1 = a1
heighta2 = a2
transmita3 = a3

Dim decompressHex() As String
ReDim decompressHex(Val("&H" & a1) * Val("&H" & a2) - 1)
Dim str1 As String, str2 As String
Dim tilenum As Long
Dim j As Long, ii As Long

Dim nn As Integer
j = nowoffset                       'j为偏移量，layer2时需要重新设置
i = 0

If a3 = "01" Then
GoTo DecompressLayer1_flag01
End If
      '第一个字节是标志位，我不知道是干什么用的，但是不能直接跳过，否则后面会少一个字节
Again:

str1 = Mid(ROMallHex, j + 1, 4)
    If Val("&H" & str1) = 0 Then
    nowoffset = j + 2
    GoTo PrintLayer01
    End If
    If (Val("&H" & str1) And Val("&H" & "8000")) <> 0 Then
    nn = Val("&H" & str1) And Val("&H" & "7FFF")
        j = j + 4
        '*******************数据写入数组
        For i = LBound(decompressHex()) To UBound(decompressHex())
        If decompressHex(i) = "" Then Exit For
        Next i
        For ii = i To i + nn - 1
        decompressHex(ii) = Mid(ROMallHex, j + 1, 2)
        Next ii
        j = j + 2
    ElseIf (Val("&H" & str1) And Val("&H" & "8000")) = 0 Then
    nn = Val("&H" & str1)
        '*******************数据写入数组
        For i = LBound(decompressHex()) To UBound(decompressHex())
        If decompressHex(i) = "" Then Exit For
        Next i
        j = j + 2
        For ii = i To i + nn - 1
        j = j + 2
        decompressHex(ii) = Mid(ROMallHex, j + 1, 2)
        Next ii
        j = j + 2
    End If

GoTo Again

DecompressLayer1_flag01:
j = 0      '用于累计总解压后的Tile个数

Dim a As Integer

Do             '解压 layer1 主循环

DoEvents
'*******************************************************这一块是解压并写入的一次循环
str1 = Mid(ROMallHex, Val(nowoffset) + 1, 2)

If Val("&H" & str1) > 128 Then               '对于大于80h的情况
tilenum = Val("&H" & str1) - 128
str2 = Mid(ROMallHex, Val(nowoffset) + 3, 2)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = str2
  Next i
j = j + tilenum
nowoffset = nowoffset + 4
Else                                        '小于等于80h
tilenum = Val("&H" & str1)
str2 = Mid(ROMallHex, Val(nowoffset) + 3, 2 * tilenum)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = Mid(str2, i * 2 - 1, 2)
  Next i
j = j + tilenum
nowoffset = nowoffset + tilenum * 2 + 2
End If
'*******************************************************写入完成，判断是否超过可用范围
If j = totlesize Then
Exit Do
ElseIf j > totlesize Then
MsgBox "overflow, decompressing failed"
Form1.Picture1.Cls
Form1.Text1.Enabled = True
Exit Sub
errorhandle:
MsgBox "Wrong!!"
End If
Loop

'*******************************************************判断完成
PrintLayer01:

Form1.Picture1.Print ""
Form1.Picture1.Print "                  layer 1:"
Form1.Picture1.Print ""

For i = 1 To Val("&H" & a1)  '写行标
Form1.Picture1.Print Right("00" & Hex(i - 1), 2) & " ";
Next i
Form1.Picture1.Print ""
Form1.Picture1.Print ""

For j = 0 To Val("&H" & a2) - 1  '写列
    For i = 0 To Val("&H" & a1) - 1  '写行
    Form1.Picture1.Print decompressHex(i + j * Val("&H" & a1)) & " ";
    Next i
Form1.Picture1.Print "       " & Hex(j)
Next j

'*******************************************************   layer1 解压数据储存到程序
For j = 0 To Val("&H" & a2) - 1 '写列
    For i = 0 To Val("&H" & a1) - 1 '写行
    Hexstream1 = Hexstream1 & decompressHex(i + j * Val("&H" & a1))
    Next i
Next j
layer1compressdatalength = nowoffset - 6
Form1.Picture1.Print "layer1 compressed data's length(byte):" & str(layer1compressdatalength / 2) & "      layer end Offset:" & Hex(Val("&H" & Form1.Text1.Text) + layer1compressdatalength / 2 + 3)
'******************************************************Sub_注出改变Room的事件地址变量提前申明
Dim MessageStream As String      ', j As Long
Dim checkStream As String
Dim H1 As String, V1 As String, H2 As String, V2 As String, TempHV As String
Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, Tempn5 As Long
Dim jj As Integer, GotoRoomID As String, GotoRoomID2 As String, GotoRoomPosition As String
Dim UsedLineTop() As Integer
ii = 0

Form1.Label2.Caption = "information：decompress finish! start making room change wireframe"

'&&&&&&&&&&&&&&&&&&&&&&&&&&&公共填充层修改&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
If Form1.Text1.Text = "3F2263" Then
MsgBox " 3F2263 cannot change，you should make a new map", vbOKOnly + vbInformation, "提示"
Form1.Text1.Enabled = True
BeforeLine = BeforeLine + 9 + 14
IfisNewRoom = True
MDIForm1.mnuroomchange.Enabled = True
Exit Sub
End If
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
IfisNewRoom = False

If Mid(ROMallHex, nowoffset + 1, 4) = "0002" Then
j = 0
nowoffset = nowoffset + 4
Do
'*******************************************************这一块是解压并写入的一次循环
str1 = Mid(ROMallHex, nowoffset + 1, 4)
If CLng("&H" & str1) > 32768 Then               '对于大于8000h的情况
tilenum = CLng("&H" & str1) - 32768
str2 = Mid(ROMallHex, nowoffset + 5, 2)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = str2
  Next i
j = j + tilenum
nowoffset = nowoffset + 6
ElseIf str1 = "0000" Then
Exit Do
Else                                        '小于等于8000h
tilenum = CLng("&H" & str1)
str2 = Mid(ROMallHex, nowoffset + 5, 2 * tilenum)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = Mid(str2, i * 2 - 1, 2)
  Next i
j = j + tilenum
nowoffset = nowoffset + tilenum * 2 + 4
End If
Loop

For j = 0 To Val("&H" & a2) - 1 '写列
    For i = 0 To Val("&H" & a1) - 1 '写行
    Hexstream2 = Hexstream2 & decompressHex(i + j * Val("&H" & a1))
    Next i
Next j
nowoffset = nowoffset + 4

layer2compressdatalength = nowoffset - 6 - layer1compressdatalength
  For i = 0 To Len(ROMallHex)
    j = i * 2
    If Mid(ROMallHex, Val(nowoffset) + 1 + 2 * i, 2) <> "00" Then
    Exit For
    End If
  Next i
leftzerozero1 = j / 2

Form1.Picture1.Print ""
Form1.Picture1.Print ""
Form1.Picture1.Print "                  layer 2:"
Form1.Picture1.Print ""

For i = 1 To Val("&H" & a1)  '写行标
Form1.Picture1.Print Right("00" & Hex(i - 1), 2) & " ";
Next i
Form1.Picture1.Print ""
Form1.Picture1.Print ""

For j = 0 To Val("&H" & a2) - 1   '写列
    For i = 0 To Val("&H" & a1) - 1   '写行
    Form1.Picture1.Print decompressHex(i + j * Val("&H" & a1)) & " ";
    Next i
Form1.Picture1.Print "       " & Hex(j)
Next j

Erase decompressHex()
Erase ROMallbyte()

'******************************************************注出改变Room的事件地址
MessageStream = ReadFileHex(gbafilepath, LevelChangeRoomStreamOffset, Right("0000" & Hex(Val("&H" & LevelChangeRoomStreamOffset) + 512), 8))
ReDim UsedLineTop(600)
For ii = 0 To 6
UsedLineTop(ii) = ii
Next ii
ii = 0
rectangleNext:

checkStream = Mid(MessageStream, ii * 24 + 1, 24)
If checkStream = "000000000000000000000000" Then
    MDIForm1.mnuroomchange.Enabled = True
    Form1.Text1.Enabled = True
    GoTo EndRectangle
End If
If Mid(checkStream, 3, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
    H1 = Mid(checkStream, 5, 2)
    V1 = Mid(checkStream, 9, 2)
    H2 = Mid(checkStream, 7, 2)
    V2 = Mid(checkStream, 11, 2)

    n1 = Val("&H" + V1) * TextHeight("FF ")
    n2 = Val("&H" + H1) * TextWidth("FF ")
    n3 = Val("&H" + V2) * TextHeight("FF ")
    n4 = Val("&H" + H2) * TextWidth("FF ")
    
    '******************************************比大小然后交换
    If n1 > n3 Then
    Tempn5 = n1
    n1 = n3
    n3 = Tempn5
    End If
    If n2 > n4 Then
    Tempn5 = n2
    n2 = n4
    n4 = Tempn5
    End If
    If Val("&H" & H1) > Val("&H" & H2) Then
    TempHV = H1
    H1 = H2
    H2 = TempHV
    End If
    If Val("&H" & V1) > Val("&H" & V2) Then
    TempHV = V1
    V1 = V2
    V2 = TempHV
    End If
    
    If Mid(checkStream, 1, 2) = "01" Then            'portal
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbBlue, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbBlue, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbBlue
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbBlue, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbBlue, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbBlue
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    ElseIf Mid(checkStream, 1, 2) = "02" Then        'vertical block
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbRed, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbRed, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbRed
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbRed, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbRed, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbRed
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    ElseIf Mid(checkStream, 1, 2) = "03" Then        'horizontal block
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbGreen, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbGreen, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbGreen
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbGreen, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbGreen, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbGreen
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    Else
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbYellow, B
    End If
NextStream:
DoEvents
ii = ii + 1    'i dont think there is more than 50 change position can be made
    If ii = 50 Then GoTo EndRectangle
GoTo rectangleNext
Else
DoEvents
ii = ii + 1    'i dont think there is more than 50 change position can be made
    If ii = 50 Then GoTo EndRectangle
GoTo rectangleNext
End If
EndRectangle:
Form1.Picture1.CurrentY = (BeforeLine + 8 + Val("&H" & heighta2)) * TextHeight("FF")
Form1.Picture1.CurrentX = 0
Form1.Picture1.ForeColor = RGB(250, 50, 250)
Form1.Picture1.Print "Room change event Block ====> Blue rectangle: protal or door   Red rectangle: immediate change room block    Green rectangle: change room with event or destination block"
Form1.Picture1.CurrentY = (BeforeLine + 8 + 7 + 2 * Val("&H" & heighta2)) * TextHeight("FF")
Form1.Picture1.ForeColor = vbBlack
'******************************************************
MDIForm1.mnuroomchange.Enabled = True
Form1.Text1.Enabled = True
BeforeLine = BeforeLine + 8 + 7 + 2 * Val("&H" & a2)
Form1.Label2.Caption = "information：finish All ! You can change Map by Ctrl + R"
Exit Sub
Else          '存在layer2

MessageStream = ReadFileHex(gbafilepath, LevelChangeRoomStreamOffset, Right("0000" & Hex(Val("&H" & LevelChangeRoomStreamOffset) + 512), 8))
ReDim UsedLineTop(600)
For ii = 0 To 6
UsedLineTop(ii) = ii
Next ii
ii = 0
rectangleNext2:

checkStream = Mid(MessageStream, ii * 24 + 1, 24)
If checkStream = "000000000000000000000000" Then
    GoTo EndRectangle2
End If
If Mid(checkStream, 3, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
    H1 = Mid(checkStream, 5, 2)
    V1 = Mid(checkStream, 9, 2)
    H2 = Mid(checkStream, 7, 2)
    V2 = Mid(checkStream, 11, 2)

    n1 = Val("&H" + V1) * TextHeight("FF ")
    n2 = Val("&H" + H1) * TextWidth("FF ")
    n3 = Val("&H" + V2) * TextHeight("FF ")
    n4 = Val("&H" + H2) * TextWidth("FF ")
    
    '******************************************比大小然后交换
    If n1 > n3 Then
    Tempn5 = n1
    n1 = n3
    n3 = Tempn5
    End If
    If n2 > n4 Then
    Tempn5 = n2
    n2 = n4
    n4 = Tempn5
    End If
    If Val("&H" & H1) > Val("&H" & H2) Then
    TempHV = H1
    H1 = H2
    H2 = TempHV
    End If
    If Val("&H" & V1) > Val("&H" & V2) Then
    TempHV = V1
    V1 = V2
    V2 = TempHV
    End If
    
    If Mid(checkStream, 1, 2) = "01" Then            'portal
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbBlue, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbBlue, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbBlue
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbBlue, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbBlue, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbBlue
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    ElseIf Mid(checkStream, 1, 2) = "02" Then        'vertical block
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbRed, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbRed, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbRed
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbRed, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbRed, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbRed
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    ElseIf Mid(checkStream, 1, 2) = "03" Then        'horizontal block
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbGreen, B
        If UsedLineTop(6 + Val("&H" & V1)) = 0 Then
        UsedLineTop(6 + Val("&H" & V1)) = 6 + Val("&H" & V1)
        Form1.Picture1.Line (n4 + TextWidth("FF"), n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), n1 + 5 + (BeforeLine + 7) * TextHeight("FF")), vbGreen, B
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = n1 + 5 + (BeforeLine + 7) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbGreen
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        Else
        For jj = 1 To 450
        If UsedLineTop(jj) = 0 Then Exit For
        Next jj
        Form1.Picture1.Line (n4 + TextWidth("FF") - 5, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF")), vbGreen, B
        Form1.Picture1.Line (n4 + TextWidth("FF"), (BeforeLine + jj) * TextHeight("FF"))-(n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF "), (BeforeLine + jj) * TextHeight("FF")), vbGreen, B
        UsedLineTop(jj) = jj
        '))))))))))))))))))))))))))))))))
        Form1.Picture1.CurrentY = (BeforeLine + jj) * TextHeight("FF")
        Form1.Picture1.CurrentX = n4 + (Val("&H" & widtha1) - Val("&H" & H1) + 4) * TextWidth("FF ")
        Form1.Picture1.ForeColor = vbGreen
        GotoRoomID = Mid(checkStream, 13, 2)
        GotoRoomID2 = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 3, 2)
        GotoRoomPosition = Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 5, 2) & Mid(MessageStream, Val("&H" & GotoRoomID) * 24 + 9, 2)
        GotoRoomID2 = Hex(Val("&H" & GotoRoomID2) + 1)
        Form1.Picture1.Print "Go to Room: "; GotoRoomID2 & "  Show Position:" & GotoRoomPosition
        Form1.Picture1.ForeColor = vbBlack
        '))))))))))))))))))))))))))))))))
        End If
    Else
        Form1.Picture1.Line (n2, n1 + (BeforeLine + 7) * TextHeight("FF"))-(n4 + TextWidth("FF"), n3 + (BeforeLine + 8) * TextHeight("FF")), vbYellow, B
    End If
NextStream2:
DoEvents
ii = ii + 1    'i dont think there is more than 50 change position can be made
    If ii = 50 Then GoTo EndRectangle2
GoTo rectangleNext2
Else
DoEvents
ii = ii + 1    'i dont think there is more than 50 change position can be made
    If ii = 50 Then GoTo EndRectangle2
GoTo rectangleNext2
End If
EndRectangle2:
Form1.Picture1.CurrentY = (BeforeLine + 8 + Val("&H" & heighta2)) * TextHeight("FF")
Form1.Picture1.CurrentX = 0
Form1.Picture1.ForeColor = RGB(250, 50, 250)
Form1.Picture1.Print "Room change event Block ====> Blue rectangle: protal or door   Red rectangle: immediate change room block    Green rectangle: change room with event or destination block"
Form1.Picture1.ForeColor = vbBlack
'******************************************************
BeforeLine = BeforeLine + 16 + 2 * Val("&H" & a2)
End If

'然后是layer 2
j = 0      '用于累计总解压后的Tile个数
a = 0
ReDim decompressHex(Val("&H" & a1) * Val("&H" & a2) - 1)
nowoffset = nowoffset + 4

                                             '对于解压标志位为02的类型在第二层的解压方法和第一类解压相同
Do             '解压 layer2 主循环

DoEvents
'*******************************************************这一块是解压并写入的二次循环
str1 = Mid(ROMallHex, Val(nowoffset) + 1, 2)

If Val("&H" & str1) > 128 Then               '对于大于80h的情况
tilenum = Val("&H" & str1) - 128
str2 = Mid(ROMallHex, Val(nowoffset) + 3, 2)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = str2
  Next i
j = j + tilenum
nowoffset = nowoffset + 4
Else                                        '小于等于80h
tilenum = Val("&H" & str1)
str2 = Mid(ROMallHex, Val(nowoffset) + 3, 2 * tilenum)
  For i = 1 To tilenum
  decompressHex(i + j - 1) = Mid(str2, i * 2 - 1, 2)
  Next i
j = j + tilenum
nowoffset = nowoffset + tilenum * 2 + 2
End If
'*******************************************************写入完成，判断是否超过可用范围
If j = totlesize Then
Exit Do
ElseIf j > totlesize Then
MsgBox "overflow, but continue decompress for success in decompress layer 1 "
Exit Do
End If
'*******************************************************判断完成
Loop

PrintLayer02:
Form1.Picture1.Print ""
Form1.Picture1.Print "                  layer 2:"
Form1.Picture1.Print ""

For i = 1 To Val("&H" & a1)  '写行标
Form1.Picture1.Print Right("00" & Hex(i - 1), 2) & " ";
Next i
Form1.Picture1.Print ""
Form1.Picture1.Print ""

For j = 0 To Val("&H" & a2) - 1   '写列
    For i = 0 To Val("&H" & a1) - 1   '写行
    Form1.Picture1.Print decompressHex(i + j * Val("&H" & a1)) & " ";
    Next i
Form1.Picture1.Print "       " & Hex(j)
Next j

'*******************************************************   layer2 解压数据储存到程序
For j = 0 To Val("&H" & a2) - 1 '写列
    For i = 0 To Val("&H" & a1) - 1 '写行
    Hexstream2 = Hexstream2 & decompressHex(i + j * Val("&H" & a1))
    Next i
Next j

  For i = 0 To Len(ROMallHex)      'a not accurate value, just to say the stored data cannot be less than "00" data in general
    j = i * 2
    If Mid(ROMallHex, Val(nowoffset) + 1 + 2 * i, 2) <> "00" Then
    Exit For
    End If
  Next i
leftzerozero1 = j / 2
layer2compressdatalength = nowoffset - layer1compressdatalength - 4 - 6 + j

Form1.Picture1.Print "layer2 compressed data's length(byte):" & str(layer2compressdatalength / 2) & "      layer end Offset:" & Hex(Val("&H" & Form1.Text1.Text) + layer1compressdatalength / 2 + layer2compressdatalength / 2 + 3 + 2)
Form1.Picture1.Print ""
Form1.Label2.Caption = "information：finish All ! You can change Map by Ctrl + R"

Erase decompressHex()                     ' release RAM
Erase ROMallbyte()
MDIForm1.mnuroomchange.Enabled = True
Form1.Text1.Enabled = True
End Sub

Private Sub Command2_Click()
Form1.Picture1.Cls
BeforeLine = 0
Form1.Label2.Caption = "informaiton:"
End Sub

Private Sub Command3_Click()
Text1.Text = "00"
End Sub

Private Sub Form_Activate()
Form1.Text1.FontSize = 15
Form1.Label1.FontSize = 15
Form1.Label2.FontSize = 15
Form1.Move 4650, 0, 18500, 11535
Load Form9
End Sub

Private Sub Form_GotFocus()
Form1.Visible = False
End Sub

Private Sub Form_Load()
Form1.Visible = False
End Sub

Private Sub HScroll1_Change()
Form1.Picture1.Left = -HScroll1.Value
DoEvents
End Sub

Private Sub Text1_Keypress(KeyCode As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
    If KeyCode = 13 Then Command1_Click
KeyCode = 0 'All other letters are unwanted.
End If
End Sub

Private Sub VScroll1_Change()
If VScroll1.Value >= 17000 Then VScroll1.Value = 17000
Form1.Picture1.Top = 375 - VScroll1.Value
Form1.Picture1.height = 12000 + VScroll1.Value
DoEvents
End Sub

