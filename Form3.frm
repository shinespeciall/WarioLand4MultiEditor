VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Rooms' connection"
   ClientHeight    =   5775
   ClientLeft      =   315
   ClientTop       =   870
   ClientWidth     =   5295
   DrawMode        =   1  'Blackness
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5295
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "save in another place(not recommand)"
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Undo All"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Change (add more to the botton)"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label2 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Room connection:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Form3Text1TextTemp As String

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_LINEFROMCHAR = &HC9


Private Sub Text1_Click()
Dim LineNumberinForm3Text1 As Long
LineNumberinForm3Text1 = LineNo(Form3.Text1.hwnd) - 1
Form3.Label2.Caption = "Line in Hex：" & Hex(LineNumberinForm3Text1)
End Sub

Function LineNo(ByVal txthwnd As Long) As Long
LineNo = SendMessageLong(txthwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
LineNo = Format$(LineNo, "##.###")
End Function

Private Sub Command1_Click()
If gbafilepath = "" Then Exit Sub
Dim strtext As String
strtext = Form3.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Dim i As Integer, j As Long
Dim str1 As String
Dim maxnum As Long
If Len(Form3.Text1.Text) = 0 Then Exit Sub
If SaveDataOffset(100) <> "" Then
    MsgBox "buffer memory used up, save all and retry !"
    Exit Sub
End If
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDatabuffer(i) = strtext Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
    End If
Next i
SaveDataOffset(i) = LevelChangeRoomStreamOffset
SaveDatabuffer(i) = strtext

For i = 0 To (Len(strtext) / 24 - 1)
For j = 0 To (Len(Form3Text1TextTemp) / 24 - 1)
If strcmp(Mid(strtext, 8 * i + 1, 24), Mid(Form3Text1TextTemp, 8 * j + 1, 24)) > 0 Then
Form9.Text1.Text = Form9.Text1.Text & "推荐测试使用的金手指：03000025:" & Mid(strtext, 8 * i + 1, 2) & vbCrLf & "测试重置的或新生成的转换点" & vbCrLf
End If
Next j
Next i

IfisNewRoomConnectionDataBuffer = True
RoomConnectionDataBuffer = strtext
End Sub

Private Sub Command2_Click()
Form3.Text1.Text = ""
If Len(gbafilepath) = 0 Then Exit Sub

Dim MessageStream As String, i As Long, j As Long
Dim checkStream As String
MessageStream = ReadFileHex(gbafilepath, LevelChangeRoomStreamOffset, Right("0000" & Hex(Val("&H" & LevelChangeRoomStreamOffset) + 1024), 8))
For i = 0 To 50     'i dont think there is more than 50 change position can be made
checkStream = Mid(MessageStream, i * 24 + 1, 24)
    If checkStream = "000000000000000000000000" Then
    Form3.Text1.Text = Form3.Text1.Text & "00 00 00 00 00 00 00 00 00 00 00 00"
    GoTo Button1Enable
    End If
For j = 1 To 12
Form3.Text1.Text = Form3.Text1.Text & Mid(checkStream, 2 * j - 1, 2) & " "
Next j
Form3.Text1.Text = Form3.Text1.Text & vbCrLf
Next i
Button1Enable:
Dim strtext As String
strtext = Form3.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")
Form3Text1TextTemp = strtext
Form3.Command1.Enabled = True
Form3.Command3.Enabled = True
End Sub

Private Sub Command3_Click()
If gbafilepath = "" Then Exit Sub
Dim strtext As String
strtext = Form3.Text1.Text
strtext = Replace(strtext, Chr(32), "")
strtext = Replace(strtext, Chr(13), "")
strtext = Replace(strtext, Chr(10), "")

Dim i As Integer, j As Long, ModValue As Integer
Dim str1 As String
Dim maxnum As Long
If SaveDataOffset(99) <> "" Then
    MsgBox "buffer memory used up, save all and retry !"
    Exit Sub
End If
For i = 1 To 100
    If SaveDataOffset(i) = "" Then Exit For
    If SaveDatabuffer(i) = strtext Then
    MsgBox "在该地址已存在保存记录", vbOKOnly + vbInformation, "警告"
    Exit Sub
    End If
str1 = ""
Next i
For j = 1 To Len(Form3Text1TextTemp) / 24
str1 = str1 & "000000000000000000000000"
Next j
SaveDataOffset(i) = LevelChangeRoomStreamOffset
SaveDatabuffer(i) = str1
ModValue = 4 - Val("&H" & SaveDatabuffer(0)) Mod 4          '如果不做字节与0，4，8，C对齐则音乐读不出来，我也不知道为什么
SaveDataOffset(i + 1) = Hex(Val("&H" & SaveDatabuffer(0)) - Val("&H" & SaveDatabuffer(0)) Mod 4 + 4)
SaveDatabuffer(i + 1) = strtext & "FFFF"          'add two Byte FF in case afterwards searching ignoring the reserved 12 "00" Byte
SaveDataOffset(i + 2) = LevelChangeRoomStreamPointerOffset
str1 = SaveDataOffset(i + 1)
str1 = Right("0" & Hex(Val("&H" & str1) + Val("&H" & "8000000")), 8)
str1 = Mid(str1, 7, 2) & Mid(str1, 5, 2) & Mid(str1, 3, 2) & Mid(str1, 1, 2)
SaveDatabuffer(i + 2) = str1
SaveDatabuffer(0) = Right("0000" & Hex(Val("&H" & SaveDatabuffer(0)) + Len(strtext) / 2 + 2 + ModValue), 8)   ' add 2 and add ModValue

For i = 0 To (Len(strtext) / 24 - 1)
For j = 0 To (Len(Form3Text1TextTemp) / 24 - 1)
If strcmp(Mid(strtext, 8 * i + 1, 24), Mid(Form3Text1TextTemp, 8 * j + 1, 24)) > 0 Then
Form9.Text1.Text = Form9.Text1.Text & "推荐测试使用的金手指：03000025:" & Mid(strtext, 8 * i + 1, 2) & vbCrLf & "测试重置的或新生成的转换点" & vbCrLf
End If
Next j
Next i

IfisNewRoomConnectionDataBuffer = True
End Sub

Private Sub Form_Activate()
Form3.Move 2000, 2000, 5530, 6360
Form3.Label1.FontSize = 15
Form3.Text1.Text = ""
Form3.Command1.Enabled = False
Form3.Command3.Enabled = False
Form3.Icon = LoadResPicture(101, vbResIcon)

If gbafilepath = "" Then Exit Sub
Command2_Click
End Sub

Private Sub Form_Load()
Form3.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form3.Text1.width = Form3.width - 400
Form3.Text1.height = Form3.height - 360 - 800 - 6360 + 3870
Form3.Command1.Top = Form3.Text1.height + Form3.Text1.Top + 100
Form3.Command3.Top = Form3.Text1.height + Form3.Text1.Top + 100
End Sub
