VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "WL4 MultiEditor"
   ClientHeight    =   6480
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16275
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuopenfile 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnunewmap 
         Caption         =   "new map"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCounterChange 
         Caption         =   "Level count down and other"
      End
      Begin VB.Menu mnudecompress 
         Caption         =   "Room decompress"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuroomchange 
         Caption         =   "Room textmap change"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRoomElement 
         Caption         =   "Room property(enimies etc.)"
      End
      Begin VB.Menu mnuRoomConnectionBeta 
         Caption         =   "Room connection info"
      End
   End
   Begin VB.Menu mnufrm 
      Caption         =   "Form"
      Begin VB.Menu mnuLevelguidefrm 
         Caption         =   "Level guide"
      End
      Begin VB.Menu mnuOutputSpritesTiles 
         Caption         =   "enimies bitmap editor"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "tools"
      Begin VB.Menu mnuFindBaseOffset 
         Caption         =   "search pointer"
      End
      Begin VB.Menu mnuCheckLoadPropertyID 
         Caption         =   "enimy flag string"
      End
      Begin VB.Menu mnuFindPropertyTableID 
         Caption         =   "find enimy flag"
      End
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "Load"
      Begin VB.Menu mnuLoadMOD 
         Caption         =   "Load MOD"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Form1.Visible = False
Form2.Visible = False
Form3.Visible = False
Form4.Visible = True
Form9.Visible = True
'Form11.Visible = True
'MDIForm1.Left = 0
'MDIForm1.Top = 0
'MDIForm1.Height = Screen.Height - 650
'MDIForm1.Width = Screen.Width

MDIForm1.mnuedit.Enabled = False
MDIForm1.mnusave.Enabled = False
MDIForm1.mnuroomchange.Enabled = False
MDIForm1.mnuFindBaseOffset.Enabled = False

WasCameraControlStringChange = False
IfisNewRoom = False
IsDeliver = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Erase SaveDatabuffer()
Erase SaveDataOffset()
End
End Sub

Private Sub mnuCheckLoadPropertyID_Click()
Dim LoadPropertyTableID As String    'pointer table after 78EF78
Dim strtxt As String, i As Integer
LoadPropertyTableID = InputBox("input Property Table ID", "searching", 0)
LoadPropertyTableID = ReadFileHex(gbafilepath, Hex(Val("&H" & LoadPropertyTableID) * 4 + Val("&H" & "78EF78")), Hex(Val("&H" & LoadPropertyTableID) * 4 + Val("&H" & "78EF78") + 3))
LoadPropertyTableID = Mid(LoadPropertyTableID, 7, 2) & Mid(LoadPropertyTableID, 5, 2) & Mid(LoadPropertyTableID, 3, 2) & Mid(LoadPropertyTableID, 1, 2)
LoadPropertyTableID = Hex(Val("&H" & LoadPropertyTableID) - Val("&H" & "8000000"))
strtxt = ReadFileHex(gbafilepath, LoadPropertyTableID, Hex(Val("&H" & LoadPropertyTableID) + 128))
For i = 0 To 30
If Mid(strtxt, 4 * i + 1, 4) = "0000" Then Exit For
Next
Do
MsgBox Mid(strtxt, 4 * i + 1, 4)
i = i - 1
Loop Until i = -1
End Sub

Private Sub mnuCounterChange_Click()
Load Form8
Form8.Visible = True
End Sub

Private Sub mnudecompress_Click()
Form1.Visible = True
End Sub

Private Sub mnuFindBaseOffset_Click()
Dim inputBaseOffset As String, inputBaseOffset2 As String

inputBaseOffset = InputBox("input start offset", "search", 0)    '___________

inputBaseOffset2 = inputBaseOffset
inputBaseOffset = Right("00000000" & CStr(inputBaseOffset), 8)
Dim i As Long, j As Long, k As Long, l As Integer
Dim stepoffset As Integer

stepoffset = Val(InputBox("step?", "search", 1))   '___________
j = Val(InputBox("input number", "search", 1))

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim PointerFirstByte As String
For k = 0 To j - 1

inputBaseOffset = Right("0" & Hex(Val("&H" & inputBaseOffset2) + k * stepoffset), 8)
inputBaseOffset = Mid(inputBaseOffset, 7, 2) & Mid(inputBaseOffset, 5, 2) & Mid(inputBaseOffset, 3, 2) & Mid(inputBaseOffset, 1, 2)
PointerFirstByte = Mid(inputBaseOffset, 1, 2)
For i = 0 To Val("&H" & "78F970") - 1
DoEvents
If Right("0" & Hex(ROMallbyte(i)), 2) = PointerFirstByte Then
If Right("0" & Hex(ROMallbyte(i)), 2) & Right("0" & Hex(ROMallbyte(i + 1)), 2) & Right("0" & Hex(ROMallbyte(i + 2)), 2) & Right("0" & Hex(ROMallbyte(i + 3)), 2) = inputBaseOffset Then
l = MsgBox("Find " & Right("0" & Hex(Val("&H" & inputBaseOffset2) + stepoffset * k), 8), vbOKCancel)
If l = 2 Then l = MsgBox("yes for next，no for break", vbYesNo)
If l = vbYes Then
    MsgBox "Offset:" & Hex(i)
    GoTo exit_i_for
ElseIf l = vbNo Then
    MsgBox "Offset:" & Hex(i)
    GoTo exit_k_for
End If
MsgBox "Offset:" & Hex(i)
End If
End If
Next i
exit_i_for:
i = 0
Next k
exit_k_for:
ROMAllHEX1 = ""
Erase ROMallbyte()
MsgBox "Finish，if no other message box has shown, no soch pointer been found.", vbOKOnly, "search pointer"
End Sub

Private Sub mnuFindPropertyTableID_Click()
Dim LoadPropertyTableID As String    'pointer table after 78EF78
Dim strtxt As String, i As Integer, j As Integer, FindStr As String
FindStr = InputBox("input enimy's Porperty flag(more than 10)", "search", 11)
For j = 0 To 89
LoadPropertyTableID = ReadFileHex(gbafilepath, Hex(j * 4 + Val("&H" & "78EF78")), Hex(j * 4 + Val("&H" & "78EF78") + 3))
LoadPropertyTableID = Mid(LoadPropertyTableID, 7, 2) & Mid(LoadPropertyTableID, 5, 2) & Mid(LoadPropertyTableID, 3, 2) & Mid(LoadPropertyTableID, 1, 2)
LoadPropertyTableID = Hex(Val("&H" & LoadPropertyTableID) - Val("&H" & "8000000"))
strtxt = ReadFileHex(gbafilepath, LoadPropertyTableID, Hex(Val("&H" & LoadPropertyTableID) + 128))
For i = 0 To 30
If Mid(strtxt, 4 * i + 1, 4) = "0000" Then Exit For
Next i
Do
If Mid(strtxt, 4 * i + 1, 2) = FindStr Then
MsgBox "Find Table index in Hex:" & Hex(j)
Exit Do
End If
i = i - 1
Loop Until i = -1
Next j
MsgBox "Finish！"
End Sub

Private Sub mnuLevelguidefrm_Click()
Load Form4
Form4.Show
End Sub

Private Sub mnuLoadMOD_Click()
CommonDialog1.Filter = "MOD File (*.txt)|*.txt"
CommonDialog1.FilterIndex = 0
CommonDialog1.CancelError = True ' 设置“CancelError”为 True
On Error Resume Next

CommonDialog1.ShowOpen

MODfilepath = CommonDialog1.FileName
Form9.Text1.Text = Form9.Text1.Text & "Load MOD File, now you can make room visually!!" & vbCrLf
If MODfilepath = "" Then
MsgBox "no file loaded", vbCritical, "Info"
Exit Sub
End If
End Sub

Private Sub mnunewmap_Click()
IfisNewRoom = True
Form2.Visible = True
End Sub

Private Sub mnuopenfile_Click()
CommonDialog1.Filter = "GBAROM File (*.gba)|*.gba"
CommonDialog1.FilterIndex = 0
CommonDialog1.CancelError = True ' 设置“CancelError”为 True
On Error Resume Next

CommonDialog1.ShowOpen

gbafilepath = CommonDialog1.FileName

If gbafilepath = "" Then
MsgBox "no file loaded", vbCritical, "Info"
Exit Sub
End If

MDIForm1.Caption = gbafilepath & "——Open"
MDIForm1.mnuedit.Enabled = True
MDIForm1.mnuFindBaseOffset.Enabled = True
MDIForm1.mnuRoomConnectionBeta.Enabled = True
Form4.Text1.Enabled = True
MDIForm1.mnusave.Enabled = True
ReDim SaveDatabuffer(100)     '记录条数最大值为101
ReDim SaveDataOffset(100)

Dim offset_78F970 As String
offset_78F970 = "78F970"

Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim nowoffset As Long    '记录偏移地址
Dim ROMallHex As String

Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, Val("&H" & offset_78F970) + 1, ROMallbyte  '从指定地址打开，第一字节的地址是1，不是0
Close #1
Dim i As Long         '转换Hex
Dim bytenum As Long '若有错误可以重新定义总读取长度
nowoffset = 0
bytenum = 4
For i = LBound(ROMallbyte) To LBound(ROMallbyte) + CLng(bytenum) - 1
ROMallHex = ROMallHex & Right("00" & Hex(ROMallbyte(i)), 2)    '用Right()防止出现"0X"的情况
Next i
SaveDataOffset(0) = offset_78F970  '统一用Hex，初始地址为0，最后储存时再加1，然后转成Dec
If ROMallHex = "FFFFFFFF" Then
SaveDatabuffer(0) = "0078F980"
Else
SaveDatabuffer(0) = ROMallHex
End If

Load Form4
Form4.Visible = True

Form4TextBox2Temp = ""
Form4.List1.Clear
Form4.List2.Clear
Form4.List3.Clear
Form4.List4.Clear
Form4.List5.Clear
Form4.List6.Clear
Form4.Text1.Text = ""
Form4.Text2.Text = ""
Form1.Text1.Text = ""
Form1.Picture1.Cls
BeforeLine = 0
WasCameraControlStringChange = False
End Sub

Private Sub mnuOutputSpritesTiles_Click()
Load Form7
Form7.Visible = True
End Sub

Private Sub mnuroomchange_Click()
Load Form2
Form2.Show
Form2.Visible = True
'IfisNewRoom = False
End Sub

Private Sub mnuRoomConnectionBeta_Click()
Load Form3
Form3.Show
Form3.Visible = True
End Sub

Private Sub mnuRoomElement_Click()
Load Form6
Form6.Visible = True
End Sub

Private Sub mnusave_Click()
Load Form5
Form5.Show
Form5.SetFocus
End Sub
