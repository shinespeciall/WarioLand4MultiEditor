VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Visual MAP Editor (Beta)"
   ClientHeight    =   12450
   ClientLeft      =   4770
   ClientTop       =   1170
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12450
   ScaleWidth      =   19305
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Left            =   18600
      Top             =   3600
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Make a model"
      Height          =   615
      Left            =   18480
      TabIndex        =   38
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Save All"
      Height          =   615
      Left            =   18480
      TabIndex        =   37
      Top             =   7080
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "layer 2"
      Height          =   375
      Left            =   15000
      TabIndex        =   36
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "layer 1"
      Height          =   615
      Left            =   15000
      TabIndex        =   35
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "layer 0"
      Height          =   495
      Left            =   15000
      TabIndex        =   34
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   11400
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Return to hex Editor"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   18480
      TabIndex        =   32
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Camera Control"
      Height          =   2175
      Left            =   360
      TabIndex        =   29
      Top             =   9120
      Width           =   3975
      Begin VB.CommandButton Command14 
         Caption         =   "Undo All"
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Add New control"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "refresh"
      Height          =   375
      Left            =   16440
      TabIndex        =   27
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "refresh with grid"
      Height          =   375
      Left            =   16440
      TabIndex        =   26
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "refresh with camera control"
      Height          =   375
      Left            =   16440
      TabIndex        =   25
      ToolTipText     =   "If no camera control it will be simply refresh"
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   3375
      Left            =   240
      ScaleHeight     =   3315
      ScaleWidth      =   4065
      TabIndex        =   24
      ToolTipText     =   "click to disable one Tile16"
      Top             =   5520
      Width           =   4125
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear"
      Height          =   495
      Left            =   18480
      TabIndex        =   23
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Undo"
      Height          =   375
      Left            =   18480
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save High Bytes"
      Height          =   495
      Left            =   13080
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Low Bytes"
      Height          =   495
      Left            =   13080
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "High Bytes"
      Height          =   495
      Left            =   11400
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Low Bytes"
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   240
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   10635
      Left            =   4560
      ScaleHeight     =   10575
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   1680
      Width           =   13635
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   390
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   780
         Left            =   7320
         Top             =   2280
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   3840
         Left            =   600
         Top             =   1320
         Visible         =   0   'False
         Width           =   5760
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MAP Properties"
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command17 
         Caption         =   "Tile size"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         ToolTipText     =   "You can render Tiles in different size"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Go"
         Height          =   495
         Left            =   2640
         TabIndex        =   21
         Top             =   3720
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "Form10.frx":0000
         Left            =   240
         List            =   "Form10.frx":0002
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "00  Debug room"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Text            =   "10"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Text            =   "10"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Text            =   "0"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Text            =   "0"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load All"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "24"
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "Present Tile Size"
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "(   ,   )"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "select MOD:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   3960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label9 
         Caption         =   "Height(Hex)"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Width(Hex)"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "vertical"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Horizontal"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "upper left show position in Hex"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Tilesets Index:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MapLength As String
Public MapHeight As String

Public MouseX As Long
Public MouseY As Long

Public Xshift As Long
Public Yshift As Long

Public WasCameraControlChange As Boolean

Public IsMakingCameraRec As Boolean            ' Used in Timer
Public IsClick As Boolean
Public WillBeResize As Integer

Public IsLayer0Change As Boolean
Public IsLayer1Change As Boolean
Public IsLayer2Change As Boolean

Private Sub Check1_Click()
Form10.Command16.Enabled = False
Form10.Picture1.Enabled = False
End Sub

Private Sub Check2_Click()
Form10.Command16.Enabled = False
Form10.Picture1.Enabled = False
End Sub

Private Sub Check3_Click()
Form10.Command16.Enabled = False
Form10.Picture1.Enabled = False
End Sub

Private Sub Combo1_Click()
Form10.Picture2.Cls
Dim Width As Integer, Height As Integer, i As Integer, j As Integer, result As Boolean
Width = Val("&" & Mid$(TileMOD(1, Form10.Combo1.ListIndex), 1, 2))
Height = Val("&" & Mid$(TileMOD(1, Form10.Combo1.ListIndex), 3, 2))
ReDim NowTileMOD(Width, Height)
For j = 0 To Height - 1
For i = 0 To Width - 1
result = DrawTile16(i, j, Mid$(TileMOD(1, Form10.Combo1.ListIndex), (j * Width + i) * 4 + 1 + 4, 4), Form10.Picture2, , DotSize)
NowTileMOD(i, j) = Mid$(TileMOD(1, Form10.Combo1.ListIndex), (j * Width + i) * 4 + 1 + 4, 4)
Next i
Next j
For i = 0 To Width
NowTileMOD(i, Height) = "0000"
Next i
For j = 0 To Height
NowTileMOD(Width, j) = "0000"
Next j
End Sub

Private Sub Command1_Click()
Dim str1 As String, i As Integer, j As Integer
If IsDeliver = True Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L0_LB_000(i, j), 3, 2)
Next i
Next j
ElseIf WholeRoomChange = True Then
If Form10.Check1.Value = 1 And Form10.Check2.Value = 0 And Form10.Check3.Value = 0 Then
For j = 0 To Layer0Height - 1
For i = 0 To Layer0Width - 1
str1 = str1 & Mid$(L0_LB_000(i, j), 3, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 1 And Form10.Check3.Value = 0 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L1_LB_000(i, j), 3, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 0 And Form10.Check3.Value = 1 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L2_LB_000(i, j), 3, 2)
Next i
Next j
Else
MsgBox "You can only choose one checkbox !", vbOKOnly, "Info"
Exit Sub
End If
End If
Form10.Text1.Text = str1
End Sub

Private Sub Command10_Click()
Form10.Check1.Enabled = False
Form10.Check2.Enabled = False
Form10.Check3.Enabled = False
Form10.Command11.Enabled = False
Form10.Command9.Enabled = False
Form10.Command10.Enabled = False
Form10.Picture1.Cls
Dim i As Integer, j As Integer, result As Boolean
If IsDeliver = True Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
ElseIf WholeRoomChange = True Then
    Dim k As Integer
    For k = 2 To 0 Step -1
    If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Layer0Height - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Layer0Width - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L2_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    End If
    Next k
End If
For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
Form10.Picture1.Line (16 * DotSize * i, 16 * DotSize * j)-(16 * DotSize * i + 16 * DotSize, 16 * DotSize * j + 16 * DotSize), vbWhite, B
DoEvents
Next i
Next j
If IsDeliver = False Then
Form10.Check1.Enabled = True
Form10.Check2.Enabled = True
Form10.Check3.Enabled = True
End If
Form10.Command11.Enabled = True
Form10.Command9.Enabled = True
Form10.Command10.Enabled = True
Form10.Command16.Enabled = True
Form10.Picture1.Enabled = True
End Sub

Private Sub Command11_Click()
Form10.Check1.Enabled = False
Form10.Check2.Enabled = False
Form10.Check3.Enabled = False
Form10.Command11.Enabled = False
Form10.Command9.Enabled = False
Form10.Command10.Enabled = False
Form10.Picture1.Cls
Dim i As Integer, j As Integer, result As Boolean
If IsDeliver = True Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    DrawTile16 i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize
    DoEvents
    Next i
    Next j
ElseIf WholeRoomChange = True Then
    Dim k As Integer
    For k = 2 To 0 Step -1
    If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Layer0Height - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Layer0Width - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    DrawTile16 i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    DrawTile16 i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    DrawTile16 i, j, L2_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize
    DoEvents
    Next i
    Next j
    End If
    Next k
End If
If IsDeliver = False Then
Form10.Check1.Enabled = True
Form10.Check2.Enabled = True
Form10.Check3.Enabled = True
End If
Form10.Command11.Enabled = True
Form10.Command9.Enabled = True
Form10.Command10.Enabled = True
Form10.Command16.Enabled = True
Form10.Picture1.Enabled = True
End Sub

Private Sub Command12_Click()
WasCameraControlChange = True
IsMakingCameraRec = True
WillBeResize = MsgBox("Make a rectangle for camera control," & vbCrLf & "Resize mode?", vbYesNo, "Info")
Form10.Shape3.Left = 0: Form10.Shape3.Top = 0
Form10.Command12.Enabled = False: Form10.Command14.Enabled = False
Form10.Shape2.Width = 780: Form10.Shape2.Height = 780
Form10.Shape2.Visible = True
Form10.Shape3.Visible = True
Form10.Timer1.Interval = 5
End Sub

Private Sub Command13_Click()
Dim i As Integer, j As Integer, IsHexstream2NeedWrite As Boolean, str1 As String
heighta2 = MapHeight
widtha1 = MapLength
Hexstream1 = ""
Hexstream2 = ""
For j = 0 To Val("&H" & heighta2) - 1
For i = 0 To Val("&H" & widtha1) - 1
Hexstream1 = Hexstream1 & "00"
If Mid$(L0_LB_000(i, j), 1, 2) <> "00" Then IsHexstream2NeedWrite = True
Next i
Next j
For j = 0 To Val("&H" & heighta2) - 1
For i = 0 To Val("&H" & widtha1) - 1
Mid$(Hexstream1, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2) = Mid$(L0_LB_000(i, j), 3, 2)
If IsHexstream2NeedWrite = True Then Hexstream2 = Hexstream2 & "00"
Next i
Next j
If IsHexstream2NeedWrite = True Then
For j = 0 To Val("&H" & heighta2) - 1
For i = 0 To Val("&H" & widtha1) - 1
Mid$(Hexstream2, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2) = Mid$(L0_LB_000(i, j), 1, 2)
Next i
Next j
End If

str1 = Replace(Form10.Text2.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
If WasCameraControlChange = True And WasCameraControlStringChange = False Then      'the latter one is for global use
IsHexstream2NeedWrite = SaveCameraString(str1)           'IsHexstream2NeedWrite is reused for another thing
If IsHexstream2NeedWrite = False Then MsgBox "fail to save camera control !"
If IsHexstream2NeedWrite = True Then MsgBox "the App save camera control in temp successfully!"
End If
Form10.Visible = False
Form2.Visible = True
MDIForm1.Enabled = True
Unload Form10
End Sub

Private Sub Command14_Click()
WasCameraControlChange = False
If Len(CameraCotrolString) <> 0 Then
Form10.Text2.Text = Mid$(CameraCotrolString, 1, 4) & vbCrLf
For i = 0 To (Len(CameraCotrolString) - 4) / 18 - 1
Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 5, 10) & vbCrLf
Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 15, 8) & vbCrLf
Next i
Else
Form10.Text2.Text = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) & "00" & vbCrLf
End If
End Sub

Private Sub Command15_Click()
Dim str1 As String
str1 = Replace(Form10.Text2.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
If WasCameraControlChange = True And WasCameraControlStringChange = False Then      'the latter one is for global use
IsHexstream2NeedWrite = SaveCameraString(str1)           'IsHexstream2NeedWrite is reused for another thing
If IsHexstream2NeedWrite = False Then MsgBox "fail to save camera control !"
If IsHexstream2NeedWrite = True Then MsgBox "the App save camera control in temp successfully!"
End If

Dim i As Long, j As Long, strtmp1 As String
Dim compressData As String, DonnotCompress As Boolean
Dim SaveResult As Boolean

If SaveDataOffset(90) <> "" Then
MsgBox "Too, much temp, Save Room Failed !", vbOKOnly, "Info"
Exit Sub
End If

If WholeRoomChange = True Then
    '----------------------------------------Layer 0
    DonnotCompress = True
    If IsLayer0Change = True Then
    strtmp1 = ""
    compressData = Right("00" & Hex(Layer0Width), 2) & Right("00" & Hex(Layer0Height), 2) & "01"
    For j = 0 To Layer0Height - 1
    For i = 0 To Layer0Width - 1
    strtmp1 = strtmp1 & Mid$(L0_LB_000(i, j), 3, 2)
    If CLng("&H" & L0_LB_000(i, j)) <> 0 Then DonnotCompress = False
    Next i
    Next j
    If DonnotCompress = False Then
        compressData = compressData & CompressDataOnly(strtmp1)
        compressData = compressData & "0001"
        strtmp1 = ""
        For j = 0 To Layer0Height - 1
        For i = 0 To Layer0Width - 1
        strtmp1 = strtmp1 & Mid$(L0_LB_000(i, j), 1, 2)
        Next i
        Next j
        compressData = compressData & CompressDataOnly(strtmp1)
        compressData = compressData & "000000FF"
        SaveResult = SaveRoomCompData(Hex(CLng("&H" & LevelAllRoomPointerandDataBaseOffset) + 44 * (CLng("&H" & LevelRoomIndex) - 1) + 8), compressData, PostlayerCompDataLength(0))
    Else
        For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
        Next i
        SaveDataOffset(i) = Hex(CLng("&H" & LevelAllRoomPointerandDataBaseOffset) + 44 * (CLng("&H" & LevelRoomIndex) - 1) + 8)
        SaveDatabuffer(i) = "63223F08"
    End If
    End If
    '----------------------------------------Layer 1
    If IsLayer1Change = True Then
    strtmp1 = ""
    compressData = Right("00" & MapLength, 2) & Right("00" & MapHeight, 2)
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    strtmp1 = strtmp1 & Mid$(L1_LB_000(i, j), 3, 2)
    Next i
    Next j
    compressData = compressData & CompressDataOnly(strtmp1)
    compressData = compressData & "0001"
    strtmp1 = ""
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    strtmp1 = strtmp1 & Mid$(L1_LB_000(i, j), 1, 2)
    Next i
    Next j
    compressData = compressData & CompressDataOnly(strtmp1)
    compressData = compressData & "000000FF"
    SaveResult = SaveRoomCompData(Hex(CLng("&H" & LevelAllRoomPointerandDataBaseOffset) + 44 * (CLng("&H" & LevelRoomIndex) - 1) + 12), compressData, PostlayerCompDataLength(1))
    End If
    '----------------------------------------Layer 2
    DonnotCompress = True
    If IsLayer2Change = True Then
    strtmp1 = ""
    compressData = Right("00" & MapLength, 2) & Right("00" & MapHeight, 2) & "01"
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    strtmp1 = strtmp1 & Mid$(L2_LB_000(i, j), 3, 2)
    If CLng("&H" & L2_LB_000(i, j)) <> 0 Then DonnotCompress = False
    Next i
    Next j
    If DonnotCompress = False Then
        compressData = compressData & CompressDataOnly(strtmp1)
        compressData = compressData & "0001"
        strtmp1 = ""
        For j = 0 To Val("&H" & MapHeight) - 1
        For i = 0 To Val("&H" & MapLength) - 1
        strtmp1 = strtmp1 & Mid$(L2_LB_000(i, j), 1, 2)
        Next i
        Next j
        compressData = compressData & CompressDataOnly(strtmp1)
        compressData = compressData & "000000FF"
        SaveResult = SaveRoomCompData(Hex(CLng("&H" & LevelAllRoomPointerandDataBaseOffset) + 44 * (CLng("&H" & LevelRoomIndex) - 1) + 16), compressData, PostlayerCompDataLength(2), True)
    Else
        For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
        Next i
        SaveDataOffset(i) = Hex(CLng("&H" & LevelAllRoomPointerandDataBaseOffset) + 44 * (CLng("&H" & LevelRoomIndex) - 1) + 16)
        SaveDatabuffer(i) = "63223F08"
    End If
    End If

    Form10.Visible = False
    Unload Form10
End If
End Sub

Private Sub Command16_Click()
ReDim MODforSave(11, 16)
Dim i As Integer, j As Integer
For j = 0 To 16
For i = 0 To 11
MODforSave(i, j) = "0000"
Next i
Next j
If (Form10.Check1.Value = 1 And Form10.Check2.Value = 0 And Form10.Check3.Value = 0) Or _
   (Form10.Check1.Value = 0 And Form10.Check2.Value = 1 And Form10.Check3.Value = 0) Or _
   (Form10.Check1.Value = 0 And Form10.Check2.Value = 0 And Form10.Check3.Value = 1) Or _
    IsDeliver = True Then
                                                                    
i = MsgBox("Get MOD from MAP ?", vbYesNo, "Info")
If i <> vbYes Then GoTo StartNewMODDialog
'--------------------------------------Use A Timer to do the rest thing

IsMakingCameraRec = True   'reuse
Form10.Shape2.Width = 780: Form10.Shape2.Height = 780
Form10.Shape2.Visible = True
Form10.Timer2.Interval = 5
Form10.Command16.Enabled = False
Exit Sub
'--------------------------------------
End If

StartNewMODDialog:
Form12.Visible = True
Form10.Enabled = False
End Sub

Private Sub Command17_Click()
Dim a As Integer
a = Val(InputBox("Input size number for Tile, the default value is 24", "Info", 24))
If a < 1 Or a > Min(Form10.Picture1.Height, Form10.Picture1.Width) Then
MsgBox "illegal value !", vbOKOnly + vbCritical, "Info"
Exit Sub
End If
DotSize = a
Form10.Label4.Caption = str(a)
Form10.Picture2.Cls
Form10.Combo1.Text = ""
Command11_Click
Form10.Shape1.Width = 15 * 16 * DotSize
Form10.Shape1.Height = 10 * 16 * DotSize
Form10.Shape3.Width = DotSize
Form10.Shape3.Height = DotSize
Form10.Shape3.Visible = False
End Sub

Private Sub Command2_Click()
Dim str1 As String, i As Integer, j As Integer
If IsDeliver = True Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L0_LB_000(i, j), 1, 2)
Next i
Next j
ElseIf WholeRoomChange = True Then
If Form10.Check1.Value = 1 And Form10.Check2.Value = 0 And Form10.Check3.Value = 0 Then
For j = 0 To Layer0Height - 1
For i = 0 To Layer0Width - 1
str1 = str1 & Mid$(L0_LB_000(i, j), 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 1 And Form10.Check3.Value = 0 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L1_LB_000(i, j), 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 0 And Form10.Check3.Value = 1 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid$(L2_LB_000(i, j), 1, 2)
Next i
Next j
Else
MsgBox "You can only choose one checkbox !", vbOKOnly, "Info"
Exit Sub
End If
End If
Form10.Text1.Text = str1
End Sub

Private Sub Command3_Click()
Dim str1 As String, i As Integer, j As Integer
str1 = Replace(Form10.Text1.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
If IsDeliver = True Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L0_LB_000(i, j), 3, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
ElseIf WholeRoomChange = True Then
If Form10.Check1.Value = 1 And Form10.Check2.Value = 0 And Form10.Check3.Value = 0 Then
For j = 0 To Layer0Height - 1
For i = 0 To Layer0Width - 1
Mid$(L0_LB_000(i, j), 3, 2) = Mid$(str1, i * 2 + j * Layer0Width * 2 + 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 1 And Form10.Check3.Value = 0 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L1_LB_000(i, j), 3, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 0 And Form10.Check3.Value = 1 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L2_LB_000(i, j), 3, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
Else
MsgBox "You can only choose one checkbox !", vbOKOnly, "Info"
Exit Sub
End If
End If
End Sub

Private Sub Command4_Click()
Dim str1 As String, i As Integer, j As Integer
str1 = Replace(Form10.Text1.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
If IsDeliver = True Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L0_LB_000(i, j), 1, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
ElseIf WholeRoomChange = True Then
If Form10.Check1.Value = 1 And Form10.Check2.Value = 0 And Form10.Check3.Value = 0 Then
For j = 0 To Layer0Height - 1
For i = 0 To Layer0Width - 1
Mid$(L0_LB_000(i, j), 1, 2) = Mid$(str1, i * 2 + j * Layer0Width * 2 + 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 1 And Form10.Check3.Value = 0 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L1_LB_000(i, j), 1, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
ElseIf Form10.Check1.Value = 0 And Form10.Check2.Value = 0 And Form10.Check3.Value = 1 Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid$(L2_LB_000(i, j), 1, 2) = Mid$(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
Else
MsgBox "You can only choose one checkbox !", vbOKOnly, "Info"
Exit Sub
End If
End If
End Sub

Private Sub Command5_Click()
Dim TileOffset As String, TileLength2 As Long
Dim TextMAPDataOffset As String, paletteOffset As String

MapLength = Form10.Text8.Text
MapHeight = Form10.Text9.Text
'If Val("&H" & MapLength) * Val("&H" & MapHeight) >= Val("&H" & "FFF") Then
'Form9.Text1.Text = Form9.Text1.Text & "Map too large, stop initrial !" & vbCrLf
'MsgBox "Map too large, please change the Length and Height value!!", vbInformation, "Info"
'Exit Sub
'End If
Dim i As Long, j As Long
'initial MAP
If IsDeliver = False And WholeRoomChange = False Then
ReDim L0_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L1_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L2_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L0_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L1_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L2_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)

For j = 1 To Val("&H" & MapHeight)
For i = 1 To Val("&H" & MapLength)

L0_LB_000(i - 1, j - 1) = "0000"
L1_LB_000(i - 1, j - 1) = "0000"
L2_LB_000(i - 1, j - 1) = "0000"
L0_LB_001(i - 1, j - 1) = "0000"
L1_LB_001(i - 1, j - 1) = "0000"
L2_LB_001(i - 1, j - 1) = "0000"
Next i
Next j
End If
'end initial

Dim Tilesets As String
Dim StrTemp As String, str1 As String, str2 As String, str3 As String, str4 As String

Form9.Text1.Text = Form9.Text1.Text & "impoting pointers......" & vbCrLf
Tilesets = Mid$(Form10.Combo2.Text, 1, 2)
StrTemp = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 9 * 4 + Val("&H" & "3F2298")), Hex(Val("&H" & Tilesets) * 9 * 4 + 35 + Val("&H" & "3F2298")))

TileOffset = Mid$(StrTemp, 5, 2) & Mid$(StrTemp, 3, 2) & Mid$(StrTemp, 1, 2)
TileLength2 = Val("&H" & Mid$(StrTemp, 11, 2) & Mid$(StrTemp, 9, 2))
Form9.Text1.Text = Form9.Text1.Text & "The amount of Tiles is " & str(TileLength2 / 32) & vbCrLf
TextMAPDataOffset = Mid$(StrTemp, 45, 2) & Mid$(StrTemp, 43, 2) & Mid$(StrTemp, 41, 2)
paletteOffset = Mid$(StrTemp, 21, 2) & Mid$(StrTemp, 19, 2) & Mid$(StrTemp, 17, 2)
ReDim Palette256(16, 16)
ReDim Tile88(TileLength2 / 16 + 64)
Form9.Text1.Text = Form9.Text1.Text & "impoting 8 * 8 Tiles Data......" & vbCrLf
Dim TextMapData As String
StrTemp = ReadFileHexWithByteInterchange(gbafilepath, TileOffset, Hex(Val("&H" & TileOffset) + TileLength2))
'-------------------------------------------------------------权宜之策
str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F8098")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F8098")))
'str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F91D8")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F91D8")))             '???
For i = 0 To 15
str2 = Mid$(str1, 4 * i + 1, 2)
Mid$(str1, 4 * i + 1, 2) = Mid$(str1, 4 * i + 3, 2)
Mid$(str1, 4 * i + 3, 2) = str2
Next i
For i = 0 To 15
str2 = Mid$(str1, 4 * i + 1, 4)
str3 = ReadFileHex(gbafilepath, Hex(Val("&H" & str2) * 8 + Val("&H" & "3F7828")), Hex(Val("&H" & str2) * 8 + Val("&H" & "3F782F")))
str2 = Mid$(str3, 9, 8)
str4 = Mid$(str3, 1, 2)
str3 = Mid$(str3, 5, 2)
If str4 = "03" Or str4 = "06" Then
str3 = Hex(Val("&H" & str3) - 1)
Else
str3 = "00"
End If
str2 = Mid$(str2, 7, 2) & Mid$(str2, 5, 2) & Mid$(str2, 3, 2) & Mid$(str2, 1, 2)
str2 = Hex(Val("&H" & str2) - Val("&H8000000") + Val("&H" & str3) * 128)
str2 = ReadFileHexWithByteInterchange(gbafilepath, str2, Hex(Val("&H" & str2) + 127))
DoEvents
For j = 0 To 3
Tile88(i * 4 + j) = Mid$(str2, j * 64 + 1, 64)
Next j
Next i
Tile88(64) = Replace(Space(64), Chr(32), "0")
'-------------------------------------------------------------End  权宜之策
For i = 0 To (TileLength2 / 16) - 1
Tile88(i + 65) = Mid$(StrTemp, 64 * i + 1, 64)
DoEvents
Next i


Form9.Text1.Text = Form9.Text1.Text & "impoting and making 16 * 16 Tiles Data......" & vbCrLf
TextMapData = ReadFileHex(gbafilepath, TextMAPDataOffset, Hex(Val("&H" & TextMAPDataOffset) + 8192))
For i = 0 To 4095
StrTemp = Mid$(TextMapData, 4 * i + 1, 2)
Mid$(TextMapData, 4 * i + 1, 2) = Mid$(TextMapData, 4 * i + 3, 2)
Mid$(TextMapData, 4 * i + 3, 2) = StrTemp
DoEvents
Next i

ReDim Tile16(1024)
Dim r0 As Long, r1 As Long, r2 As Long
For i = 0 To 1023
r0 = i * 4
r2 = r0 Or 1
Tile16(i) = Mid$(TextMapData, r0 * 2 * 2 + 1, 4)
r1 = (r2 + 1) * 2 ^ 16
Tile16(i) = Tile16(i) & Mid$(TextMapData, r2 * 2 * 2 + 1, 4)
r0 = 128 * 2 ^ 9
r2 = r1 + r0
r1 = RSH(r1, 15)
Tile16(i) = Tile16(i) & Mid$(TextMapData, r1 * 2 + 1, 4)
r2 = RSH(r2, 15)
Tile16(i) = Tile16(i) & Mid$(TextMapData, r2 * 2 + 1, 4)
DoEvents
Next i

Form9.Text1.Text = Form9.Text1.Text & "impoting palette 256 ......" & vbCrLf
StrTemp = ReadFileHex(gbafilepath, paletteOffset, Hex(Val("&H" & paletteOffset) + 256 * 2 - 1))
For j = 0 To 15
For i = 0 To 15
Palette256(i, j) = RGB555ToRGB888(Mid$(StrTemp, 64 * j + 4 * i + 1, 4))
DoEvents
Next i
Next j

If IsDeliver = False And WholeRoomChange = False Then
Form9.Text1.Text = Form9.Text1.Text & "No Map Text Data inport, create new map!" & vbCrLf
Form9.Text1.Text = Form9.Text1.Text & "Making grid......" & vbCrLf
For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
Form10.Picture1.Line (64 * 6 * i, 64 * 6 * j)-(64 * 6 * i + 64 * 6, 64 * 6 * j + 64 * 6), vbWhite, B
DoEvents
Next i
Next j
End If

If IsDeliver = True Then
Form9.Text1.Text = Form9.Text1.Text & "Rendering......" & vbCrLf
DoEvents
For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
DrawTile16 i, j, L0_LB_000(i, j), Form10.Picture1, , DotSize
Next i
Next j
End If

If WholeRoomChange = True Then
Dim Flag01 As Integer
Flag01 = Val("&H" & Mid$(LevelAllRoomPointerandDataallHex, 1 + 52 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2))     'eleventh byte flag
If (Flag01 Mod 4) = 0 Then
layerPriority(1) = 1: layerPriority(2) = 2
ElseIf (Flag01 Mod 4) = 1 Then
layerPriority(0) = 1: layerPriority(2) = 2
ElseIf (Flag01 Mod 4) = 2 Then
layerPriority(0) = 1: layerPriority(2) = 2
Else
layerPriority(0) = 2: layerPriority(2) = 1
End If

Form9.Text1.Text = Form9.Text1.Text & "Rendering......" & vbCrLf
Dim k As Integer
For k = 2 To 0 Step -1
If layerPriority(0) = k Then
For j = 0 To Min(Layer0Height - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Layer0Width - 1 - Xshift, Form10.Picture1.Width \ DotSize)
DrawTile16 i, j, L0_LB_000(i, j), Form10.Picture1, , DotSize
Next i
Next j
ElseIf layerPriority(1) = k Then
For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
DrawTile16 i, j, L1_LB_000(i, j), Form10.Picture1, , DotSize
Next i
Next j
ElseIf layerPriority(2) = k Then
For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
DrawTile16 i, j, L2_LB_000(i, j), Form10.Picture1, , DotSize
Next i
Next j
End If
Next k

Form10.Check1.Value = 1
If ExistUnchangeableLayer0 = True Then
Form10.Check1.Enabled = False
Form10.Check1.Value = 0
End If
Form10.Check2.Value = 1
Form10.Check3.Value = 1
End If

Form9.Text1.Text = Form9.Text1.Text & "Finish All" & vbCrLf
Form10.Combo1.Enabled = True
Form10.Command5.Enabled = False
Form10.Command16.Enabled = True
Form10.Picture1.Enabled = True
MDIForm1.Enabled = False
End Sub


Private Sub Command6_Click()
Dim i As Integer, j As Integer, result As Boolean
If IsDeliver = True Then
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    L0_LB_000(i, j) = L0_LB_001(i, j)
    L0_LB_001(i, j) = "0000"
    Next i
    Next j
ElseIf WholeRoomChange = True Then
    If LastLayerChange = 0 Then
    For j = 0 To Layer0Height - 1
    For i = 0 To Layer0Width - 1
    L0_LB_000(i, j) = L0_LB_001(i, j)
    L0_LB_001(i, j) = "0000"
    Next i
    Next j
    ElseIf LastLayerChange = 1 Then
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    L1_LB_000(i, j) = L1_LB_001(i, j)
    L1_LB_001(i, j) = "0000"
    Next i
    Next j
    ElseIf LastLayerChange = 2 Then
    For j = 0 To Val("&H" & MapHeight) - 1
    For i = 0 To Val("&H" & MapLength) - 1
    L2_LB_000(i, j) = L2_LB_001(i, j)
    L2_LB_001(i, j) = "0000"
    Next i
    Next j
    End If
End If
Command11_Click
Form10.Command6.Enabled = False
End Sub

Private Sub Command7_Click()
Xshift = Val("&H" & Form10.Text6.Text)
Yshift = Val("&H" & Form10.Text7.Text)
Command11_Click
End Sub

Private Sub Command8_Click()
Form10.Picture1.Cls
End Sub

Private Sub Command9_Click()
Form10.Check1.Enabled = False
Form10.Check2.Enabled = False
Form10.Check3.Enabled = False
Form10.Command11.Enabled = False
Form10.Command9.Enabled = False
Form10.Command10.Enabled = False
Form10.Picture1.Cls
Dim i As Integer, j As Integer, result As Boolean
If IsDeliver = True Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
ElseIf WholeRoomChange = True Then
    Dim k As Integer
    For k = 2 To 0 Step -1
    If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Layer0Height - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Layer0Width - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L0_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
    For j = 0 To Min(Val("&H" & MapHeight) - 1 - Yshift, Form10.Picture1.Height \ DotSize)
    For i = 0 To Min(Val("&H" & MapLength) - 1 - Xshift, Form10.Picture1.Width \ DotSize)
    result = DrawTile16(i, j, L2_LB_000(i + Xshift, j + Yshift), Form10.Picture1, , DotSize)
    DoEvents
    Next i
    Next j
    End If
    Next k
End If

Dim OutputString As String, kk As Integer
OutputString = Replace(Form10.Text2.Text, Chr(32), "")
OutputString = Replace(OutputString, Chr(13), "")
OutputString = Replace(OutputString, Chr(10), "")
If Len(OutputString) <> 0 Then
    Dim b0 As Integer, b1 As Integer, b2 As Integer, b3 As Integer, b4 As Integer, b5 As Integer
            kk = Val("&H" & Mid$(OutputString, 3, 2))
            For j = 0 To (kk - 1)
            b0 = Val("&H" & Mid$(OutputString, 18 * j + 7, 2))
            b1 = Val("&H" & Mid$(OutputString, 18 * j + 9, 2))
            b2 = Val("&H" & Mid$(OutputString, 18 * j + 11, 2))
            b3 = Val("&H" & Mid$(OutputString, 18 * j + 13, 2))
            If b0 > b1 Then
            b4 = b0
            b0 = b1
            b1 = b4
            End If
            If b2 > b1 Then
            b4 = b2
            b2 = b3
            b3 = b4
            End If
            Form10.Picture1.Line ((b0 - Xshift) * DotSize * 16, (b2 - Yshift) * DotSize * 16)-((b1 + 1 - Xshift) * DotSize * 16, (b3 + 1 - Yshift) * DotSize * 16), vbRed, B                  'posibly overflow
            Form10.Picture1.Line ((b0 - Xshift) * DotSize * 16, (b2 - Yshift) * DotSize * 16)-((b0 - Xshift + 0.5) * DotSize * 16, (b2 - Yshift + 0.5) * DotSize * 16), vbGreen, BF
            Form10.Picture1.Line ((b1 - Xshift + 0.5) * DotSize * 16, (b3 - Yshift + 0.5) * DotSize * 16)-((b1 - Xshift + 1) * DotSize * 16, (b3 - Yshift + 1) * DotSize * 16), vbGreen, BF
            If Mid$(OutputString, 18 * j + 15, 2) <> "FF" Then
            b4 = Val("&H" & Mid$(OutputString, 18 * j + 15, 2))
            b5 = Val("&H" & Mid$(OutputString, 18 * j + 17, 2))
            Form10.Picture1.Line ((b4 - Xshift) * DotSize * 16, (b5 - Yshift) * DotSize * 16)-((b4 + 1 - Xshift) * DotSize * 16, (b5 + 1 - Yshift) * DotSize * 16), vbRed, B
            If Val("&H" & Mid$(OutputString, 18 * j + 19, 2)) = "00" Then b0 = Val("&H" & Mid$(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid$(OutputString, 18 * j + 19, 2)) = "01" Then b1 = Val("&H" & Mid$(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid$(OutputString, 18 * j + 19, 2)) = "02" Then b2 = Val("&H" & Mid$(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid$(OutputString, 18 * j + 19, 2)) = "03" Then b3 = Val("&H" & Mid$(OutputString, 18 * j + 21, 2))
            Form10.Picture1.Line ((b0 - Xshift) * DotSize * 16, (b2 - Yshift) * DotSize * 16)-((b1 + 1 - Xshift) * DotSize * 16, (b3 + 1 - Yshift) * DotSize * 16), vbYellow, B
            Form10.Picture1.Line ((b0 - Xshift) * DotSize * 16, (b2 - Yshift) * DotSize * 16)-((b0 - Xshift + 0.5) * DotSize * 16, (b2 - Yshift + 0.5) * DotSize * 16), vbWhite, BF
            Form10.Picture1.Line ((b1 - Xshift + 0.5) * DotSize * 16, (b3 - Yshift + 0.5) * DotSize * 16)-((b1 - Xshift + 1) * DotSize * 16, (b3 - Yshift + 1) * DotSize * 16), vbWhite, BF
            End If
            Next j
End If
If IsDeliver = False Then
Form10.Check1.Enabled = True
Form10.Check2.Enabled = True
Form10.Check3.Enabled = True
End If
Form10.Command11.Enabled = True
Form10.Command9.Enabled = True
Form10.Command10.Enabled = True
Form10.Command16.Enabled = True
Form10.Picture1.Enabled = True
End Sub

Private Sub Form_Load()
Form10.Icon = LoadResPicture(101, vbResIcon)
IsMakingCameraRec = False
IsClick = False
If MODfilepath = "" Then
MODfilepath = App.Path & "\MOD\" & Mid$(LevelAllRoomPointerandDataallHex, 1 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) & " Block.txt"
    If Dir(MODfilepath) = "" Then
    Open MODfilepath For Append As #3
    Print #3, "000 0000 For Clean"
    Print #3, "01010000"
    Print #3, "000 0040 Fill Block"
    Print #3, "01010040"
    Print #3, "000 0002 RoomChangeMark"
    Print #3, "01010002";
    Close #3
    End If
End If

Dim i As Integer, j As Integer
ReDim TileMOD(1, 500)
i = 0
Open MODfilepath For Input As #2
Do While Not EOF(2)
    Line Input #2, TileMOD(0, i)
    Form10.Combo1.AddItem TileMOD(0, i)
    Line Input #2, TileMOD(1, i)
i = i + 1
Loop
Close #2
Form10.Picture1.BackColor = &H0&
Form10.Combo1.FontSize = 15
Form10.Combo2.FontSize = 15
Form10.Combo2.AddItem "00  Debug room"
Form10.Combo2.AddItem "01  Palm Tree Paradise"
Form10.Combo2.AddItem "02  Caves"
Form10.Combo2.AddItem "03  The Big Board"
Form10.Combo2.AddItem "04  The Big Board"
Form10.Combo2.AddItem "05  The Big Board?"
Form10.Combo2.AddItem "06  Wildflower Fields"
Form10.Combo2.AddItem "07  Toy Block Tower"
Form10.Combo2.AddItem "08  Factory"
Form10.Combo2.AddItem "09  Wildflower Underground"
Form10.Combo2.AddItem "0A  Wildflower WaterPlace"
Form10.Combo2.AddItem "0B  Underwater"
Form10.Combo2.AddItem "0C  Toy Block Tower"
Form10.Combo2.AddItem "0D  Toy Block Tower"
Form10.Combo2.AddItem "0E  Toy Block Tower"
Form10.Combo2.AddItem "0F  Doodle"
Form10.Combo2.AddItem "10  Dominoes"
Form10.Combo2.AddItem "11  Hall of Hieroglyphs"
Form10.Combo2.AddItem "12  Haunte House (plus debug tiles)"
Form10.Combo2.AddItem "13  Crescent Moon Village outside"
Form10.Combo2.AddItem "14  Drain"
Form10.Combo2.AddItem "15  Arabian outside"
Form10.Combo2.AddItem "16  Arabian inside"
Form10.Combo2.AddItem "17  Arabian"
Form10.Combo2.AddItem "18  Arabian"
Form10.Combo2.AddItem "19  Arabian"
Form10.Combo2.AddItem "1A  Dominoes (blue)"
Form10.Combo2.AddItem "1B  Dominoes (purple)"
Form10.Combo2.AddItem "1C  Dominoes (teal)"
Form10.Combo2.AddItem "1D  Factory"
Form10.Combo2.AddItem "1E  Factory"
Form10.Combo2.AddItem "1F  Jungle"
Form10.Combo2.AddItem "20  Factory"
Form10.Combo2.AddItem "21  Junkyard"
Form10.Combo2.AddItem "22  Junkyard"
Form10.Combo2.AddItem "23  Pinball"
Form10.Combo2.AddItem "24  Pinball"
Form10.Combo2.AddItem "25  Pinball (with Gorilla)"
Form10.Combo2.AddItem "26  Jungle"
Form10.Combo2.AddItem "27  40 Below Fridge?"
Form10.Combo2.AddItem "28  Jungle"
Form10.Combo2.AddItem "29  Jungle caves"
Form10.Combo2.AddItem "2A  Hotel"
Form10.Combo2.AddItem "2B  Hotel"
Form10.Combo2.AddItem "2C  Hotel"
Form10.Combo2.AddItem "2D  Hotel"
Form10.Combo2.AddItem "2E  Hotel"
Form10.Combo2.AddItem "2F  Hotel (outside)"
Form10.Combo2.AddItem "30  Unused in-game (Haunted House)"
Form10.Combo2.AddItem "31  Unused in-game (Haunted House)"
Form10.Combo2.AddItem "32  Unused in-game (Cardboard)"
Form10.Combo2.AddItem "33  Cardboard"
Form10.Combo2.AddItem "34  Caves"
Form10.Combo2.AddItem "35  Jungle"
Form10.Combo2.AddItem "36  Caves"
Form10.Combo2.AddItem "37  Lava level"
Form10.Combo2.AddItem "38  Caves"
Form10.Combo2.AddItem "39  Golden Passage"
Form10.Combo2.AddItem "3A  Hotel"
Form10.Combo2.AddItem "3B  Hotel"
Form10.Combo2.AddItem "3C  Hotel"
Form10.Combo2.AddItem "3D  Hotel"
Form10.Combo2.AddItem "3E  40 Below Fridge"
Form10.Combo2.AddItem "3F  Factory"
Form10.Combo2.AddItem "40  Factory"
Form10.Combo2.AddItem "41  Arabian"
Form10.Combo2.AddItem "42  Boss room"
Form10.Combo2.AddItem "43  Boss corridor"
Form10.Combo2.AddItem "44  Boss room"
Form10.Combo2.AddItem "45  Frozen lava level"
Form10.Combo2.AddItem "46  Lava level"
Form10.Combo2.AddItem "47  Hall of Hieroglyphs"
Form10.Combo2.AddItem "48  Boss room"
Form10.Combo2.AddItem "49  Boss room"
Form10.Combo2.AddItem "4A  Boss corridor"
Form10.Combo2.AddItem "4B  Boss corridor"
Form10.Combo2.AddItem "4C  Boss corridor"
Form10.Combo2.AddItem "4D  Boss corridor"
Form10.Combo2.AddItem "4E  Boss corridor"
Form10.Combo2.AddItem "4F  Boss room (Diva)"
Form10.Combo2.AddItem "50  Hall of Hieroglyphs"
Form10.Combo2.AddItem "51  Jungle"
Form10.Combo2.AddItem "52  Wildflower"
Form10.Combo2.AddItem "53  Crescent Moon Village"
Form10.Combo2.AddItem "54  Crescent Moon Village"
Form10.Combo2.AddItem "55  Crescent Moon Village"
Form10.Combo2.AddItem "56  Toy Block Tower"
Form10.Combo2.AddItem "57  Pinball"
Form10.Combo2.AddItem "58  Bonus room"
Form10.Combo2.AddItem "59  Bonus room"
Form10.Combo2.AddItem "5A  Final level"
Form10.Combo2.AddItem "5B  The Big Board end"

Xshift = 0
Yshift = 0
WasCameraControlChange = False
DotSize = 24
Form10.Label4.Caption = str(DotSize)

If IsDeliver = True Then
    Form10.Command15.Enabled = False
    Form10.Command15.Visible = False
    ReDim L0_LB_000(Val("&H" & widtha1) - 1, Val("&H" & heighta2) - 1)
    ReDim L0_LB_001(Val("&H" & widtha1) - 1, Val("&H" & heighta2) - 1)
    For j = 1 To Val("&H" & heighta2)
    For i = 1 To Val("&H" & widtha1)
    L0_LB_000(i - 1, j - 1) = "0000"
    L0_LB_001(i - 1, j - 1) = "0000"
    Next i
    Next j
    Form10.Text8.Text = widtha1
    Form10.Text9.Text = heighta2

    For j = 0 To Val("&H" & heighta2) - 1
    For i = 0 To Val("&H" & widtha1) - 1
    Mid$(L0_LB_000(i, j), 3, 2) = Mid$(Hexstream1, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2)
    Next i
    Next j

    If Hexstream2 <> "" Then
    For j = 0 To Val("&H" & heighta2) - 1
    For i = 0 To Val("&H" & widtha1) - 1
     Mid$(L0_LB_000(i, j), 1, 2) = Mid$(Hexstream2, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2)
    Next i
    Next j
    End If

    Form10.Combo2.ListIndex = Val("&H" & Mid$(LevelAllRoomPointerandDataallHex, 1 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2))
    Form10.Command13.Enabled = True
    Form10.Check1.Enabled = False
    Form10.Check2.Enabled = False
    Form10.Check3.Enabled = False
    Command5_Click

    If Len(CameraCotrolString) <> 0 Then
    Form10.Text2.Text = Mid$(CameraCotrolString, 1, 4) & vbCrLf
    For i = 0 To (Len(CameraCotrolString) - 4) / 18 - 1
    Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 5, 10) & vbCrLf
    Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 15, 8) & vbCrLf
    Next i
    Else
    Form10.Text2.Text = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) & "00" & vbCrLf
    End If
    If WasCameraControlStringChange = True Then
    Form10.Command12.Enabled = False
    Form10.Command14.Enabled = False
    End If
ElseIf WholeRoomChange = True Then
    Form10.Combo2.ListIndex = Val("&H" & Mid$(LevelAllRoomPointerandDataallHex, 1 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2))
    Form10.Command13.Enabled = False
    Form10.Command13.Visible = False
    Form10.Text8.Text = Hex(layerWidth)
    Form10.Text9.Text = Hex(layerHeight)
    layerHeight = 0
    layerWidth = 0
    Erase TextMap()
    ReDim layerPriority(2)
    layerPriority(0) = 0
    layerPriority(1) = 0
    layerPriority(2) = 0
    LastLayerChange = 4
    Command5_Click
    
    If Len(CameraCotrolString) <> 0 Then
    Form10.Text2.Text = Mid$(CameraCotrolString, 1, 4) & vbCrLf
    For i = 0 To (Len(CameraCotrolString) - 4) / 18 - 1
    Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 5, 10) & vbCrLf
    Form10.Text2.Text = Form10.Text2.Text & Mid$(CameraCotrolString, 18 * i + 15, 8) & vbCrLf
    Next i
    Else
    Form10.Text2.Text = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) & "00" & vbCrLf
    End If
    If WasCameraControlStringChange = True Then
    Form10.Command12.Enabled = False
    Form10.Command14.Enabled = False
    End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form10.Shape1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
IsDeliver = False
WholeRoomChange = False

Erase Tile16()
Erase Tile88()
Erase Palette256()

Erase L0_LB_000()
Erase L1_LB_000()
Erase L2_LB_000()
Erase L0_LB_001()
Erase L1_LB_001()
Erase L2_LB_001()

Erase TileMOD()
Erase layerPriority()
Erase PostlayerCompDataLength()

MDIForm1.Enabled = True
MDIForm1.SetFocus
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form10.Shape1.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form10.Shape1.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form10.Shape1.Visible = True
MouseX = X \ (DotSize * 16)
MouseY = Y \ (DotSize * 16)
Form10.Label3.Caption = "(" & Hex(MouseX) & " , " & Hex(MouseY) & ")"
Form10.Shape1.Left = X - Xshift
Form10.Shape1.Top = Y - Yshift
End Sub

Private Sub Picture1_Click()
Dim i As Integer, j As Integer

If IsMakingCameraRec = False And Form10.Combo1.Text <> "" Then          'Start a new room is not support
    If IsDeliver = True Then
        For j = 0 To Val("&H" & MapHeight) - 1
        For i = 0 To Val("&H" & MapLength) - 1
        L0_LB_001(i, j) = L0_LB_000(i, j)
        Next i
        Next j
        
        For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
        For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
        If NowTileMOD(i, j) <> "0000" Then
        result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
        L0_LB_000(MouseX + Xshift + i, MouseY + Yshift + j) = NowTileMOD(i, j)
        Form10.Command6.Enabled = True
        End If
        Next i
        Next j
        
        If (UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) = 1) And (UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) = 1) And NowTileMOD(0, 0) = "0000" And NowTileMOD(1, 0) = "0000" And NowTileMOD(0, 1) = "0000" And NowTileMOD(1, 1) = "0000" Then
        result = DrawTile16(MouseX, MouseY, NowTileMOD(0, 0), Form10.Picture1)
        L0_LB_000(MouseX + Xshift, MouseY + Yshift) = NowTileMOD(0, 0)
        End If
    ElseIf WholeRoomChange = True Then
        If Form10.Check1.Value = 1 And layerPriority(0) < layerPriority(1) And layerPriority(0) < layerPriority(2) Then
        For j = 0 To Layer0Height - 1
        For i = 0 To Layer0Width - 1
        L0_LB_001(i, j) = L0_LB_000(i, j)
        Next i
        Next j
        IsLayer0Change = True
        LastLayerChange = 0
        ElseIf Form10.Check2.Value = 1 And layerPriority(1) < layerPriority(2) Then
        For j = 0 To Val("&H" & MapHeight) - 1
        For i = 0 To Val("&H" & MapLength) - 1
        L1_LB_001(i, j) = L1_LB_000(i, j)
        Next i
        Next j
        IsLayer1Change = True
        LastLayerChange = 1
        ElseIf Form10.Check2.Value = 1 Then
        For j = 0 To Val("&H" & MapHeight) - 1
        For i = 0 To Val("&H" & MapLength) - 1
        L2_LB_001(i, j) = L2_LB_000(i, j)
        Next i
        Next j
        IsLayer2Change = True
        LastLayerChange = 2
        End If
        
        Dim k As Integer
        
        If LastLayerChange = 0 Then
        For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
        For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
        If NowTileMOD(i, j) <> "0000" Then
            result = DrawTile16(MouseX + i, MouseY + j, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            End If
            Next k
            L0_LB_000(MouseX + Xshift + i, MouseY + Yshift + j) = NowTileMOD(i, j)
            Form10.Command6.Enabled = True
        End If
        Next i
        Next j

        If (UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) = 1) And (UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) = 1) And NowTileMOD(0, 0) = "0000" And NowTileMOD(1, 0) = "0000" And NowTileMOD(0, 1) = "0000" And NowTileMOD(1, 1) = "0000" Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L1_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L2_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            End If
            Next k
        L0_LB_000(MouseX + Xshift, MouseY + Yshift) = NowTileMOD(0, 0)      '"0000"
        End If

        ElseIf LastLayerChange = 1 Then
        For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
        For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
        If NowTileMOD(i, j) <> "0000" Then
            result = DrawTile16(MouseX + i, MouseY + j, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            End If
            Next k
            L1_LB_000(MouseX + Xshift + i, MouseY + Yshift + j) = NowTileMOD(i, j)
            Form10.Command6.Enabled = True
        End If
        Next i
        Next j
        
        If (UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) = 1) And (UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) = 1) And NowTileMOD(0, 0) = "0000" And NowTileMOD(1, 0) = "0000" And NowTileMOD(0, 1) = "0000" And NowTileMOD(1, 1) = "0000" Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L0_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L2_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            End If
            Next k
        L1_LB_000(MouseX + Xshift, MouseY + Yshift) = NowTileMOD(0, 0)      '"0000"
        End If

        ElseIf LastLayerChange = 2 Then
        For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
        For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
        If NowTileMOD(i, j) <> "0000" Then
            result = DrawTile16(MouseX + i, MouseY + j, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 And (Layer0Height - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 And (Val("&H" & MapHeight) - 1 - Yshift) >= 0 Then
            result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1, , DotSize)
            End If
            Next k
            L2_LB_000(MouseX + Xshift + i, MouseY + Yshift + j) = NowTileMOD(i, j)
            Form10.Command6.Enabled = True
        End If
        Next i
        Next j

        If (UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) = 1) And (UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) = 1) And NowTileMOD(0, 0) = "0000" And NowTileMOD(1, 0) = "0000" And NowTileMOD(0, 1) = "0000" And NowTileMOD(1, 1) = "0000" Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, True, DotSize)
            For k = 2 To 0 Step -1
            If layerPriority(0) = k And Form10.Check1.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L0_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            ElseIf layerPriority(1) = k And Form10.Check2.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, L1_LB_000(MouseX + Xshift, MouseY + Yshift), Form10.Picture1, , DotSize)
            ElseIf layerPriority(2) = k And Form10.Check3.Value = 1 Then
            result = DrawTile16(MouseX, MouseY, "0000", Form10.Picture1, , DotSize)
            End If
            Next k
        L2_LB_000(MouseX + Xshift, MouseY + Yshift) = NowTileMOD(0, 0)      '"0000"
        End If

        End If
    End If
Else
IsClick = True
End If
End Sub

Private Sub Picture2_Click()
Dim result As Boolean, result2 As Boolean
NowTileMOD(MouseX, MouseY) = "0000"
Dim i As Integer, j As Integer
result = True
result2 = True
For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
If NowTileMOD(0, j) <> "0000" Then result2 = False
Next j
For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
If NowTileMOD(i, 0) <> "0000" Then result = False
Next i
If result = True Then
For i = LBound(NowTileMOD, 1) To UBound(NowTileMOD, 1)
For j = LBound(NowTileMOD, 2) To UBound(NowTileMOD, 2) - 1
NowTileMOD(i, j) = NowTileMOD(i, j + 1)
Next j
NowTileMOD(i, j) = "0000"
Next i
End If
If result2 = True Then
For j = LBound(NowTileMOD, 2) To UBound(NowTileMOD, 2)
For i = LBound(NowTileMOD, 1) To UBound(NowTileMOD, 1) - 1
NowTileMOD(i, j) = NowTileMOD(i + 1, j)
Next i
NowTileMOD(i, j) = "0000"
Next j
End If
For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
result = DrawTile16(i, j, NowTileMOD(i, j), Form10.Picture2, , DotSize)
Next i
Next j
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X \ (24 * 16)
MouseY = Y \ (24 * 16)
End Sub

Private Sub Timer1_Timer()
Static a As Boolean                'start by False
Static x1 As Long, x2 As Long, y1 As Long, y2 As Long, XT As Long, YT As Long
Static OneTile As Boolean

If WillBeResize = vbYes Then
    If OneTile = False And a = False And x1 = 0 Then          'drawing part
    Form10.Shape2.Left = MouseX * (DotSize * 16)
    Form10.Shape2.Top = MouseY * (DotSize * 16)
    ElseIf OneTile = False And a = True And x1 > 0 Then
    Form10.Shape2.Width = (MouseX + 1) * (DotSize * 16) - Form10.Shape2.Left
    Form10.Shape2.Height = (MouseY + 1) * (DotSize * 16) - Form10.Shape2.Top
    ElseIf OneTile = True Then
    Form10.Shape3.Left = MouseX * (DotSize * 16)
    Form10.Shape3.Top = MouseY * (DotSize * 16)
    ElseIf OneTile = False And a = False And XT > 0 Then
    If (MouseX + Xshift) < x1 And (MouseY + Yshift) < y2 And (MouseY + Yshift) > y1 Then
    Form10.Shape2.Left = MouseX * (DotSize * 16)
    Form10.Shape2.Width = (x2 + 1) * (DotSize * 16) - Form10.Shape2.Left
    ElseIf (MouseX + Xshift) > x2 And (MouseY + Yshift) < y2 And (MouseY + Yshift) > y1 Then
    Form10.Shape2.Width = (MouseX + 1) * (DotSize * 16) - Form10.Shape2.Left
    ElseIf (MouseY + Yshift) < y1 And (MouseX + Xshift) < x2 And (MouseX + Xshift) > x1 Then
    Form10.Shape2.Top = MouseY * (DotSize * 16)
    Form10.Shape2.Height = (y2 + 1) * (DotSize * 16) - Form10.Shape2.Top
    ElseIf (MouseY + Yshift) > y2 And (MouseX + Xshift) < x2 And (MouseX + Xshift) > x1 Then
    Form10.Shape2.Height = (MouseY + 1) * (DotSize * 16) - Form10.Shape2.Top
    Else
    Form10.Shape2.Left = x1 * (DotSize * 16)
    Form10.Shape2.Width = (x2 + 1) * (DotSize * 16) - Form10.Shape2.Left
    Form10.Shape2.Top = y1 * (DotSize * 16)
    Form10.Shape2.Height = (y2 + 1) * (DotSize * 16) - Form10.Shape2.Top
    End If
    End If

    If IsClick = True And OneTile = False And y2 = 0 Then
        a = Not a
        If a = True Then
            x1 = MouseX + Xshift
            y1 = MouseY + Yshift
        Else
            x2 = MouseX + Xshift
            y2 = MouseY + Yshift
            Form10.Text2.Text = Form10.Text2.Text & "02" & Right("0" & Hex(x1), 2) & Right("0" & Hex(x2), 2) & Right("0" & Hex(y1), 2) & Right("0" & Hex(y2), 2) & vbCrLf
            OneTile = True
        End If
        IsClick = False
    ElseIf IsClick = True And OneTile = True Then
        XT = MouseX + Xshift
        YT = MouseY + Yshift
        Form10.Text2.Text = Form10.Text2.Text & Right("0" & Hex(XT), 2) & Right("0" & Hex(YT), 2)
        Form10.Shape3.Visible = False
        OneTile = False
        IsClick = False
    ElseIf IsClick = True And (MouseX + Xshift) < x1 And (MouseY + Yshift) < y2 And (MouseY + Yshift) > y1 Then
    Form10.Text2.Text = Form10.Text2.Text & "00" & Right("0" & Hex(MouseX + Xshift), 2) & vbCrLf
    a = False: OneTile = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0: XT = 0: YT = 0: Form10.Timer1.Interval = 0: Form10.Shape2.Visible = False: Form10.Command12.Enabled = True: Form10.Command14.Enabled = True: WillBeResize = 0
    IsClick = False: IsMakingCameraRec = False: Call ChangeNumberAFlag
    ElseIf IsClick = True And (MouseX + Xshift) > x2 And (MouseY + Yshift) < y2 And (MouseY + Yshift) > y1 Then
    Form10.Text2.Text = Form10.Text2.Text & "01" & Right("0" & Hex(MouseX + Xshift), 2) & vbCrLf
    a = False: OneTile = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0: XT = 0: YT = 0: Form10.Timer1.Interval = 0: Form10.Shape2.Visible = False: Form10.Command12.Enabled = True: Form10.Command14.Enabled = True: WillBeResize = 0
    IsClick = False: IsMakingCameraRec = False: Call ChangeNumberAFlag
    ElseIf IsClick = True And (MouseY + Yshift) < y1 And (MouseX + Xshift) < x2 And (MouseX + Xshift) > x1 Then
    Form10.Text2.Text = Form10.Text2.Text & "02" & Right("0" & Hex(MouseY + Yshift), 2) & vbCrLf
    a = False: OneTile = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0: XT = 0: YT = 0: Form10.Timer1.Interval = 0: Form10.Shape2.Visible = False: Form10.Command12.Enabled = True: Form10.Command14.Enabled = True: WillBeResize = 0
    IsClick = False: IsMakingCameraRec = False: Call ChangeNumberAFlag
    ElseIf IsClick = True And (MouseY + Yshift) > y2 And (MouseX + Xshift) < x2 And (MouseX + Xshift) > x1 Then
    Form10.Text2.Text = Form10.Text2.Text & "03" & Right("0" & Hex(MouseY + Yshift), 2) & vbCrLf
    a = False: OneTile = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0: XT = 0: YT = 0: Form10.Timer1.Interval = 0: Form10.Shape2.Visible = False: Form10.Command12.Enabled = True: Form10.Command14.Enabled = True: WillBeResize = 0
    IsClick = False: IsMakingCameraRec = False: Call ChangeNumberAFlag
    End If
ElseIf WillBeResize = vbNo Then
    If a = False And x1 = 0 Then          'drawing part
    Form10.Shape2.Left = MouseX * (DotSize * 16)
    Form10.Shape2.Top = MouseY * (DotSize * 16)
    ElseIf a = True And x1 > 0 Then
    Form10.Shape2.Width = (MouseX + 1) * (DotSize * 16) - Form10.Shape2.Left
    Form10.Shape2.Height = (MouseY + 1) * (DotSize * 16) - Form10.Shape2.Top
    End If
    
    If IsClick = True And y2 = 0 Then
        a = Not a
        If a = True Then
            x1 = MouseX + Xshift
            y1 = MouseY + Yshift
            IsClick = False
        Else
            x2 = MouseX + Xshift
            y2 = MouseY + Yshift
            Form10.Text2.Text = Form10.Text2.Text & "02" & Right("0" & Hex(x1), 2) & Right("0" & Hex(x2), 2) & Right("0" & Hex(y1), 2) & Right("0" & Hex(y2), 2) & vbCrLf & "FFFFFFFF" & vbCrLf
            a = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0: WillBeResize = 0
            Form10.Shape3.Visible = False: Form10.Timer1.Interval = 0: Form10.Shape2.Visible = False: Form10.Command12.Enabled = True: Form10.Command14.Enabled = True: Form10.Shape3.Visible = False
            IsClick = False: IsMakingCameraRec = False: Call ChangeNumberAFlag
        End If
    End If
End If
End Sub

Public Sub ChangeNumberAFlag()
Dim str As String, i As Integer
str = Replace(Form10.Text2.Text, Chr(32), "")
str = Replace(str, Chr(13), "")
str = Replace(str, Chr(10), "")
If Len(str) <> 0 Then
Mid$(str, 3, 4) = Right("0" & Hex((Len(str) - 4) / 18), 2)
Form10.Text2.Text = Mid$(str, 1, 4) & vbCrLf
For i = 0 To (Len(str) - 4) / 18 - 1
Form10.Text2.Text = Form10.Text2.Text & Mid$(str, 18 * i + 5, 10) & vbCrLf
Form10.Text2.Text = Form10.Text2.Text & Mid$(str, 18 * i + 15, 8) & vbCrLf
Next i
End If
End Sub

Private Sub Timer2_Timer()
Static a As Boolean                'start by False
Static x1 As Long, x2 As Long, y1 As Long, y2 As Long

If a = False And x1 = 0 Then          'drawing part
    Form10.Shape2.Left = MouseX * (DotSize * 16)
    Form10.Shape2.Top = MouseY * (DotSize * 16)
ElseIf a = True And x1 > 0 Then
    Form10.Shape2.Width = (MouseX + 1) * (DotSize * 16) - Form10.Shape2.Left
    Form10.Shape2.Height = (MouseY + 1) * (DotSize * 16) - Form10.Shape2.Top
End If

If IsClick = True And y2 = 0 Then
    a = Not a
    If a = True Then
        x1 = MouseX + Xshift
        y1 = MouseY + Yshift
        IsClick = False
    Else
        x2 = MouseX + Xshift
        y2 = MouseY + Yshift
        
        Dim i As Integer, j As Integer
        If x2 - x1 < 12 And y2 - y1 < 17 Then
        If Form10.Check1.Value = 1 Or IsDeliver = True Then
        For j = 0 To y2 - y1
        For i = 0 To x2 - x1
        MODforSave(i, j) = L0_LB_000(i + x1, j + y1)
        Next i
        Next j
        ElseIf Form10.Check2.Value = 1 Then
        For j = 0 To y2 - y1
        For i = 0 To x2 - x1
        MODforSave(i, j) = L1_LB_000(i + x1, j + y1)
        Next i
        Next j
        Else
        For j = 0 To y2 - y1
        For i = 0 To x2 - x1
        MODforSave(i, j) = L2_LB_000(i + x1, j + y1)
        Next i
        Next j
        End If
        End If
        
        a = False: x1 = 0: y1 = 0: x2 = 0: y2 = 0
        Form10.Timer2.Interval = 0: Form10.Shape2.Visible = False: Form10.Command16.Enabled = True
        IsClick = False: IsMakingCameraRec = False: Form12.Visible = True: Form10.Enabled = False
    End If
End If

End Sub
