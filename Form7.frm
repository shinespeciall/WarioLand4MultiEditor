VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   Caption         =   "Sprites and Tiles Editor"
   ClientHeight    =   5400
   ClientLeft      =   2385
   ClientTop       =   2325
   ClientWidth     =   22830
   LinkTopic       =   "Form7"
   ScaleHeight     =   5760
   ScaleMode       =   0  'User
   ScaleWidth      =   22830
   Begin MSComctlLib.Slider Slider1 
      Height          =   555
      Left            =   20880
      TabIndex        =   9
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      _Version        =   393216
      Max             =   15
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   21000
      ScaleHeight     =   315
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "change color"
      Height          =   495
      Left            =   21000
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reshow"
      Height          =   615
      Left            =   21000
      TabIndex        =   4
      ToolTipText     =   "If bitmap been erased by some reason use this to reshow"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   615
      Left            =   21000
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   20715
      TabIndex        =   2
      Top             =   120
      Width           =   20775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   21000
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   615
      Left            =   21000
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Palette adjustor"
      Height          =   375
      Left            =   21120
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "if you have offset"
      Height          =   255
      Left            =   21000
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TilesOffset As String
Public TilesLength As String
Public ColorHByte As String
Public ifMouseDowm As Boolean
Public SpritesID As String

Private Sub Command1_Click()
If Len(gbafilepath) = 0 Then Exit Sub
Form7.Picture1.Cls
Dim TileData As String
Dim i As Integer, j As Integer, k As Integer
Dim TempPointer As String, LenthMessage As String
If Len(Form7.Text1.Text) <> 0 Then GoTo PrintbyMyself
TempPointer = "78EBF0"
LenthMessage = "3B2C90"
Dim LineNum As Integer
SpritesID = InputBox("Sprites ID (more than 10 in Hex):", "Request", 1)
If Val("&H" & SpritesID) < 17 Then
MsgBox "Wrong ID !"
Exit Sub
End If
SpritesID = Hex(Val("&H" & SpritesID) - 16)
ReDim Palette16Color(16)
'----------------------------------------------Make palette Beta Mode
Dim paletteStr As String
paletteStr = ReadFileHex(gbafilepath, Hex(Val("&H78EDB4") + 4 * Val("&H" & SpritesID)), Hex(Val("&H78EDB4") + 4 * Val("&H" & SpritesID) + 2))
'paletteStr = Mid(paletteStr, 5, 2) & Mid(paletteStr, 3, 2) & Mid(paletteStr, 1, 2)    'get offset
paletteStr = Hex(Val("&H" & Mid(paletteStr, 5, 2) & Mid(paletteStr, 3, 2) & Mid(paletteStr, 1, 2)) + 32 * Form7.Slider1.Value)

paletteStr = ReadFileHex(gbafilepath, paletteStr, Hex(Val("&H" & paletteStr) + 31))

For i = 0 To 15
Palette16Color(i) = RGB555ToRGB888(Mid(paletteStr, 4 * i + 1, 4))
Next i

'----------------------------------------------End Make palette

TempPointer = Hex(Val("&H" & TempPointer) + 4 * Val("&H" & SpritesID))
TempPointer = ReadFileHex(gbafilepath, TempPointer, Hex(Val("&H" & TempPointer) + 4))
TempPointer = Mid(TempPointer, 7, 2) & Mid(TempPointer, 5, 2) & Mid(TempPointer, 3, 2) & Mid(TempPointer, 1, 2)
TempPointer = Hex(Val("&H" & TempPointer) - Val("&H" & "8000000"))
LenthMessage = Hex(Val("&H" & LenthMessage) + 4 * Val("&H" & SpritesID))
LenthMessage = ReadFileHex(gbafilepath, LenthMessage, Hex(Val("&H" & LenthMessage) + 4))
LenthMessage = Mid(LenthMessage, 7, 2) & Mid(LenthMessage, 5, 2) & Mid(LenthMessage, 3, 2) & Mid(LenthMessage, 1, 2)
TilesOffset = TempPointer
TilesLength = LenthMessage
TileData = ReadFileHexWithByteInterchange(gbafilepath, TempPointer, Hex(Val("&H " & TempPointer) + 2 * Val("&H" & LenthMessage)))
ReDim TempPointerValue(2 * Val("&H" & TilesLength) - 1)
For LineNum = 0 To Val("&H" & LenthMessage) / Val("&H" & "400") - 1
For k = 0 To 31                'horizontal tile number
For j = 0 To 7                 'vertical
For i = 0 To 7                 'count by byte, horizontal
'Form7.Picture1.Line (80 * i + 640 * k, 80 * j + 640 * LineNum)-(80 * i + 640 * k + 80, 80 * j + 80 + 640 * LineNum), QBColor(Val("&H" & Mid(TileData, i + 8 * j + 64 * k + 1, 1))), BF
Form7.Picture1.Line (80 * i + 640 * k, 80 * j + 640 * LineNum)-(80 * i + 640 * k + 80, 80 * j + 80 + 640 * LineNum), Palette16Color(Val("&H" & Mid(TileData, i + 8 * j + 64 * k + 1, 1))), BF
TempPointerValue(i + 8 * j + 64 * k + 64 * 32 * LineNum) = Mid(TileData, i + 8 * j + 64 * k + 1, 1)
DoEvents
Next i
Next j
Next k
TileData = Right(TileData, Len(TileData) - 64 * 32)
Next LineNum
Form7.Command2.Enabled = True
Form7.Command3.Enabled = True
Form7.Command4.Enabled = True
Form7.Picture1.Enabled = True
Exit Sub

PrintbyMyself:
ReDim TempPointerValue(2 * Val("&H" & TilesLength) - 1)
TempPointer = Form7.Text1.Text
TilesOffset = TempPointer
Form7.Text1.Text = ""
LenthMessage = InputBox("输入Tile加载长度，单位是双字，数值以16进制表示", "加载Tile", 10)
TilesLength = LenthMessage
TileData = ReadFileHexWithByteInterchange(gbafilepath, TempPointer, Hex(Val("&H " & TempPointer) + 2 * Val("&H" & LenthMessage)))
For LineNum = 0 To Val("&H" & LenthMessage) / Val("&H" & "400") - 1
For k = 0 To 31                'horizontal tile number
For j = 0 To 7                 'vertical
For i = 0 To 7                 'count by byte, horizontal
Form7.Picture1.Line (80 * i + 640 * k, 80 * j + 640 * LineNum)-(80 * i + 640 * k + 80, 80 * j + 80 + 640 * LineNum), QBColor(Val("&H" & Mid(TileData, i + 8 * j + 64 * k + 1, 1))), BF
TempPointerValue(i + 8 * j + 64 * k + 64 * 32 * LineNum) = Mid(TileData, i + 8 * j + 64 * k + 1, 1)
DoEvents
Next i
Next j
Next k
TileData = Right(TileData, Len(TileData) - 64 * 32)
Next LineNum
Form7.Command2.Enabled = True
Form7.Command3.Enabled = True
Form7.Command4.Enabled = True
Form7.Picture1.Enabled = True
Form7.Slider1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim i As Long, jj As Long, j As String

For i = LBound(TempPointerValue()) + 1 To UBound(TempPointerValue()) Step 2
j = TempPointerValue(i - 1)
TempPointerValue(i - 1) = TempPointerValue(i)
TempPointerValue(i) = j
Next i

    If SaveDataOffset(99) <> "" Then
        MsgBox "buffer memory used up, save all and retry !"
        Exit Sub
    End If
    For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
    Next i
    
SaveDataOffset(i) = TilesOffset
Dim strtxt As String
For jj = LBound(TempPointerValue()) To UBound(TempPointerValue())
strtxt = strtxt & TempPointerValue(jj)
Next jj
SaveDatabuffer(i) = strtxt

MsgBox "录入完成！"
Unload Form7
End Sub

Private Sub Command3_Click()
Form7.Picture1.Cls
Dim i As Integer, j As Integer, k As Integer
Dim LineNum As Integer
For LineNum = 0 To Val("&H" & TilesLength) / Val("&H" & "400") - 1
For k = 0 To 31                'horizontal tile number
For j = 0 To 7                 'vertical
For i = 0 To 7                 'count by byte, horizontal
Form7.Picture1.Line (80 * i + 640 * k, 80 * j + 640 * LineNum)-(80 * i + 640 * k + 80, 80 * j + 80 + 640 * LineNum), Palette16Color(Val("&H" & TempPointerValue(i + 8 * j + 64 * k + 64 * 32 * LineNum))), BF
DoEvents
Next i
Next j
Next k
Next LineNum
End Sub

Private Sub Command4_Click()
ColorHByte = Right(InputBox("input a halfByte from 0 to F for choose one color, 0 is the background color which will be transparent", "Info", 1), 1)
'Form7.Picture2.BackColor = QBColor(Val("&H" & ColorHByte))
Form7.Picture2.BackColor = Palette16Color(Val("&H" & ColorHByte))
End Sub

Private Sub Form_Activate()
Form7.Text1.Text = ""
Form7.Text1.FontSize = 12
ifMouseDowm = False

Form7.Command2.Enabled = False
Form7.Command3.Enabled = False
Form7.Command4.Enabled = False
Form7.Picture1.Enabled = False
Form7.Slider1.Enabled = False

MDIForm1.Enabled = False
End Sub

Private Sub Form_Resize()                 ' 80 twip(缇 读ti，第二声) per height and width
If Form7.height < 1700 Then Exit Sub
Form7.Picture1.Move 60, 60, 20560, (Form7.height \ 640 - 1) * 640 'height: 640 per 2 line   Width: 20480
ColorHByte = "1"
If TileLength <> "" Then Form7.Picture2.BackColor = Palette16Color(Val("&H" & ColorHByte))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase Palette16Color()
MDIForm1.Enabled = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ifMouseDowm = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If ifMouseDowm = True Then
Form7.Picture1.Line ((X \ 80) * 80, (Y \ 80) * 80)-((X \ 80) * 80 + 80, (Y \ 80) * 80 + 80), Palette16Color(Val("&H" & ColorHByte)), BF

Dim i As Integer, j As Integer, k As Integer
Dim LineNum As Integer
LineNum = Y \ 640
k = X \ 640
i = (X Mod 640) \ 80
j = (Y Mod 640) \ 80
TempPointerValue(i + 8 * j + 64 * k + 64 * 32 * LineNum) = ColorHByte
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ifMouseDowm = False
End Sub

Private Sub Slider1_Change()
Dim paletteStr As String, i As Integer
paletteStr = ReadFileHex(gbafilepath, Hex(Val("&H78EDB4") + 4 * Val("&H" & SpritesID)), Hex(Val("&H78EDB4") + 4 * Val("&H" & SpritesID) + 2))
'paletteStr = Mid(paletteStr, 5, 2) & Mid(paletteStr, 3, 2) & Mid(paletteStr, 1, 2)    'get offset
paletteStr = Hex(Val("&H" & Mid(paletteStr, 5, 2) & Mid(paletteStr, 3, 2) & Mid(paletteStr, 1, 2)) + 32 * Form7.Slider1.Value)
paletteStr = ReadFileHex(gbafilepath, paletteStr, Hex(Val("&H" & paletteStr) + 31))
For i = 0 To 15
Palette16Color(i) = RGB555ToRGB888(Mid(paletteStr, 4 * i + 1, 4))
Next i
Command3_Click
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
