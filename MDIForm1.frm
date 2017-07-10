VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "WL4 MultiEditor"
   ClientHeight    =   6480
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16275
   Icon            =   "MDIForm1.frx":0000
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
      Caption         =   "Function"
      Begin VB.Menu mnuLevelguidefrm 
         Caption         =   "Level guide"
      End
      Begin VB.Menu mnuOutputSpritesTiles 
         Caption         =   "enimies bitmap editor"
      End
      Begin VB.Menu mnuoutput 
         Caption         =   "Output"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "tools"
      Begin VB.Menu mnuFindBaseOffset 
         Caption         =   "search pointer"
      End
      Begin VB.Menu mnuCheckLoadPropertyID 
         Caption         =   "Enumerate a Sprites Table"
      End
      Begin VB.Menu mnuFindPropertyTableID 
         Caption         =   "Find Sprites in table"
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
MDIForm1.Icon = LoadResPicture(101, vbResIcon)

MDIForm1.mnuedit.Enabled = False
MDIForm1.mnusave.Enabled = False
MDIForm1.mnuroomchange.Enabled = False
MDIForm1.mnuFindBaseOffset.Enabled = False
MDIForm1.mnuLevelguidefrm.Enabled = False
MDIForm1.mnuOutputSpritesTiles.Enabled = False
MDIForm1.mnuCheckLoadPropertyID.Enabled = False
MDIForm1.mnuFindPropertyTableID.Enabled = False
MDIForm1.mnuCounterChange.Enabled = False
MDIForm1.mnuRoomConnectionBeta.Enabled = False
MDIForm1.mnudecompress.Enabled = False
MDIForm1.mnuRoomElement.Enabled = False

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
Dim strtxt As String, j As Integer
LoadPropertyTableID = InputBox("input Sprites Table ID", "searching", 0)
LoadPropertyTableID = ReadFileHex(gbafilepath, Hex(Val("&H" & LoadPropertyTableID) * 4 + Val("&H" & "78EF78")), Hex(Val("&H" & LoadPropertyTableID) * 4 + Val("&H" & "78EF78") + 3))
LoadPropertyTableID = Mid(LoadPropertyTableID, 7, 2) & Mid(LoadPropertyTableID, 5, 2) & Mid(LoadPropertyTableID, 3, 2) & Mid(LoadPropertyTableID, 1, 2)
LoadPropertyTableID = Hex(Val("&H" & LoadPropertyTableID) - Val("&H" & "8000000"))
strtxt = ReadFileHex(gbafilepath, LoadPropertyTableID, Hex(Val("&H" & LoadPropertyTableID) + 128))
For j = 0 To 30
If Mid(strtxt, 4 * j + 1, 4) = "0000" Then Exit For
Form9.Text1.Text = Form9.Text1.Text & Mid(strtxt, 4 * j + 1, 4) & " "
Next j
Form9.Text1.Text = Form9.Text1.Text & vbCrLf
Form9.Text1.Text = Form9.Text1.Text & "Finish !   PS: the first byte is the Sprites ID"
Form9.Text1.Text = Form9.Text1.Text & vbCrLf
End Sub

Private Sub mnuCounterChange_Click()
Load Form8
Form8.Visible = True
End Sub

Private Sub mnudecompress_Click()
Form1.Visible = True
End Sub

Private Sub mnuFindBaseOffset_Click()
If gbafilepath = "" Then
Form9.Text1.Text = Form9.Text1.Text & "You haven't open a gba file !" & vbCrLf
Exit Sub
End If
Dim inputBaseOffset As String, inputBaseOffset2 As String
inputBaseOffset = InputBox("input start pointer", "search", 8000000)
If inputBaseOffset = "" Then Exit Sub
inputBaseOffset2 = inputBaseOffset
inputBaseOffset = Right("00000000" & CStr(inputBaseOffset), 8)
Dim i As Long, j As Long, k As Long
Dim stepoffset As Integer
stepoffset = Val(InputBox("step width? (In Dec)", "search", 1))
If stepoffset = 0 Then Exit Sub
j = Val(InputBox("input number of steps, count from one", "search", 1))
If j <= 0 Then Exit Sub
Dim ROMallbyte() As Byte     'max ROM space is 32 MB, is in VB's changeable String Type, its maximun is 2^31
Dim ROMallHex As String
Form9.Text1.Text = Form9.Text1.Text & "Start searching ! Please wait..." & vbCrLf
Open gbafilepath For Binary As #1
ReDim ROMallbyte(LOF(1) - 1)
Get #1, , ROMallbyte   'ROMallstr now contains all of the text in the file
Close #1
Dim PointerFirstByte As String
For k = 0 To j
inputBaseOffset = Right("0" & Hex(Val("&H" & inputBaseOffset2) + k * stepoffset), 8)
inputBaseOffset = Mid(inputBaseOffset, 7, 2) & Mid(inputBaseOffset, 5, 2) & Mid(inputBaseOffset, 3, 2) & Mid(inputBaseOffset, 1, 2)
PointerFirstByte = Mid(inputBaseOffset, 1, 2)
For i = 0 To Val("&H" & "78F970") - 1
DoEvents
If Right("0" & Hex(ROMallbyte(i)), 2) = PointerFirstByte Then
If Right("0" & Hex(ROMallbyte(i)), 2) & Right("0" & Hex(ROMallbyte(i + 1)), 2) & Right("0" & Hex(ROMallbyte(i + 2)), 2) & Right("0" & Hex(ROMallbyte(i + 3)), 2) = inputBaseOffset Then
Form9.Text1.Text = Form9.Text1.Text & "Find " & Right("0" & Hex(Val("&H" & inputBaseOffset2) + stepoffset * k), 8)
Form9.Text1.Text = Form9.Text1.Text & " Offset: " & Hex(i) & vbCrLf
End If
End If
Next i
Next k
ROMAllHEX1 = ""
Erase ROMallbyte()
Form9.Text1.Text = Form9.Text1.Text & "Finish，if no message has shown, no such pointer been found." & vbCrLf
End Sub

Private Sub mnuFindPropertyTableID_Click()
Dim LoadPropertyTableID As String    'pointer table after 78EF78
Dim strtxt As String, i As Integer, j As Integer, FindStr As String
FindStr = InputBox("input Sprite's ID (more than 10 in Hex)", "search", 11)
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
Form9.Text1.Text = Form9.Text1.Text & "Table ID (Hex):" & Hex(j) & vbCrLf
Exit Do
End If
i = i - 1
Loop Until i = -1
Next j
Form9.Text1.Text = Form9.Text1.Text & "Finish！" & vbCrLf
End Sub

Private Sub mnuLevelguidefrm_Click()
Load Form4
Form4.Show
End Sub

Private Sub mnuLoadMOD_Click()
CommonDialog1.Filter = "MOD File (*.txt)|*.txt"
CommonDialog1.FilterIndex = 0
CommonDialog1.CancelError = False ' 设置“CancelError”为 False
CommonDialog1.FileName = ""
On Error Resume Next

CommonDialog1.ShowOpen

MODfilepath = CommonDialog1.FileName
If MODfilepath = "" Then
MsgBox "no file loaded", vbCritical, "Info"
Exit Sub
End If
Load Form9
Form9.Text1.Text = ""
If MODfilepath <> "" Then Form9.Text1.Text = "Load MOD File, now you can make room visually!!" & vbCrLf
End Sub

Private Sub mnunewmap_Click()
IfisNewRoom = True
Form2.Visible = True
End Sub

Private Sub mnuopenfile_Click()
CommonDialog1.Filter = "GBAROM File (*.gba)|*.gba"
CommonDialog1.FilterIndex = 0
CommonDialog1.CancelError = False ' 设置“CancelError”为 False
CommonDialog1.FileName = ""
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
MDIForm1.mnuOutputSpritesTiles.Enabled = True
MDIForm1.mnuCheckLoadPropertyID.Enabled = True
MDIForm1.mnuFindPropertyTableID.Enabled = True
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
WasCameraControlStringChange = False

Dim k As Long, str5 As String                  ' Initialize form4 combo1
For k = 0 To 23
str0 = ""
str5 = GetLevelNamePointer(k)
str5 = Mid(str5, 7, 2) & Mid(str5, 5, 2) & Mid(str5, 3, 2) & Mid(str5, 1, 2)
str5 = Hex(Val("&H" & str5) - Val("&H8000000"))
str5 = ReadFileHex(gbafilepath, str5, Hex(Val("&H" & str5) + 25))
For i = 0 To 25
str0 = str0 & DEX_to_letter(CLng(Val("&H" & Mid(str5, 2 * i + 1, 2))))
Next i
Form4.Combo1.AddItem Right("00" & Hex(k), 2) & str0
Next k
MDIForm1.mnuLevelguidefrm.Enabled = True
MDIForm1.mnuopenfile.Enabled = False
End Sub

Private Sub mnuoutput_Click()
Load Form9
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
