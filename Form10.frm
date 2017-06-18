VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Visual MAP Editor"
   ClientHeight    =   15000
   ClientLeft      =   495
   ClientTop       =   465
   ClientWidth     =   24645
   LinkTopic       =   "Form1"
   ScaleHeight     =   15000
   ScaleWidth      =   24645
   Visible         =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "refresh"
      Height          =   375
      Left            =   21840
      TabIndex        =   30
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      Caption         =   "refresh with grid"
      Height          =   375
      Left            =   21840
      TabIndex        =   29
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "refresh with camera control"
      Height          =   375
      Left            =   21840
      TabIndex        =   28
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3195
      ScaleWidth      =   3705
      TabIndex        =   27
      ToolTipText     =   "click to disable one Tile16"
      Top             =   6840
      Width           =   3765
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear"
      Height          =   615
      Left            =   23880
      TabIndex        =   26
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Undo"
      Height          =   615
      Left            =   23880
      TabIndex        =   22
      Top             =   2640
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Caption         =   "to layer 3"
      Height          =   255
      Left            =   20400
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "to layer 2"
      Height          =   255
      Left            =   20400
      TabIndex        =   20
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "to layer 1"
      Height          =   255
      Left            =   20400
      TabIndex        =   19
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save High Bytes"
      Height          =   495
      Left            =   18480
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Low Bytes"
      Height          =   495
      Left            =   18480
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "High Bytes"
      Height          =   495
      Left            =   16800
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Low Bytes"
      Height          =   495
      Left            =   16800
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
      Width           =   12135
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   13035
      Left            =   4560
      ScaleHeight     =   12975
      ScaleWidth      =   19095
      TabIndex        =   1
      Top             =   1680
      Width           =   19155
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   3840
         Left            =   3600
         Top             =   3120
         Width           =   5760
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame"
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command7 
         Caption         =   "Go"
         Height          =   495
         Left            =   3120
         TabIndex        =   24
         Top             =   3960
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "Form10.frx":0000
         Left            =   240
         List            =   "Form10.frx":0002
         TabIndex        =   23
         Text            =   "00  Debug room"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2040
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
         Left            =   1800
         TabIndex        =   13
         Text            =   "0"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Text            =   "0"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load All"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "select MOD:"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   3960
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label9 
         Caption         =   "Height(Hex)"
         Height          =   375
         Left            =   2040
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
         Left            =   600
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Horizontal"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "upper left show position in Hex"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   3240
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

Private Sub Combo1_Click()
Form10.Picture2.Cls
Dim width As Integer, height As Integer, i As Integer, j As Integer, result As Boolean
width = Val(Mid(TileMOD(1, Form10.Combo1.ListIndex), 1, 2))
height = Val(Mid(TileMOD(1, Form10.Combo1.ListIndex), 3, 2))
ReDim NowTileMOD(width, height)
For j = 0 To height - 1
For i = 0 To width - 1
result = DrawTile16(i, j, Mid(TileMOD(1, Form10.Combo1.ListIndex), (j * width + i) * 4 + 1 + 4, 4), Form10.Picture2)
NowTileMOD(i, j) = Mid(TileMOD(1, Form10.Combo1.ListIndex), (j * width + i) * 4 + 1 + 4, 4)
Next i
Next j
For i = 0 To width
NowTileMOD(i, height) = "0000"
Next i
For j = 0 To height
NowTileMOD(width, j) = "0000"
Next j
End Sub

Private Sub Command1_Click()
Dim str1 As String, i As Integer, j As Integer
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid(L1_LB_000(i, j), 3, 2)
Next i
Next j
Form10.Text1.Text = str1
End Sub

Private Sub Command10_Click()
Form10.Picture1.Cls
Form10.Picture1.DrawWidth = 1
Dim i As Integer, j As Integer, result As Boolean
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1)
DoEvents
Next i
Next j
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
Form10.Picture1.Line (64 * 6 * i, 64 * 6 * j)-(64 * 6 * i + 64 * 6, 64 * 6 * j + 64 * 6), vbWhite, B
DoEvents
Next i
Next j
End Sub

Private Sub Command11_Click()
Form10.Picture1.Cls
Form10.Picture1.DrawWidth = 2
Dim i As Integer, j As Integer, result As Boolean
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1)
DoEvents
Next i
Next j
End Sub

Private Sub Command2_Click()
Dim str1 As String, i As Integer, j As Integer
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
str1 = str1 & Mid(L1_LB_000(i, j), 1, 2)
Next i
Next j
Form10.Text1.Text = str1
End Sub

Private Sub Command3_Click()
Dim str1 As String, i As Integer, j As Integer
str1 = Replace(Form10.Text1.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Mid(L1_LB_000(i, j), 3, 2) = Mid(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
End Sub

Private Sub Command4_Click()
Dim str1 As String, i As Integer, j As Integer
str1 = Replace(Form10.Text1.Text, Chr(32), "")
str1 = Replace(str1, Chr(13), "")
str1 = Replace(str1, Chr(10), "")
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
 Mid(L1_LB_000(i, j), 1, 2) = Mid(str1, i * 2 + j * Val("&H" & MapLength) * 2 + 1, 2)
Next i
Next j
End Sub

Private Sub Command5_Click()
'Public TileOffset As String, TileLength2 As Long
'Public TextMAPDataOffset As String, paletteOffset As String
If gbafilepath = "" Then
Form9.Text1.Text = Form9.Text1.Text & "Please Open the GBA file!!" & vbCrLf
Exit Sub
End If

MapLength = Form10.Text8.Text
MapHeight = Form10.Text9.Text
If Val("&H" & MapLength) * Val("&H" & MapHeight) >= Val("&H" & "FFF") Then
Form9.Text1.Text = Form9.Text1.Text & "Map too large, stop initrial !" & vbCrLf
MsgBox "Map too large, please change the Length and Height value!!", vbInformation, "Info"
Exit Sub
End If
Dim i As Long, j As Long
'initial MAP matrix
If IsDeliver <> True Then
ReDim L1_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L2_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L3_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L1_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L2_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L3_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)

For j = 1 To Val("&H" & MapHeight)
For i = 1 To Val("&H" & MapLength)
L1_LB_000(i - 1, j - 1) = "0000"
L2_LB_000(i - 1, j - 1) = "0000"
L3_LB_000(i - 1, j - 1) = "0000"
L1_LB_001(i - 1, j - 1) = "0000"
L2_LB_001(i - 1, j - 1) = "0000"
L3_LB_001(i - 1, j - 1) = "0000"
Next i
Next j
End If

If IsDeliver <> True Then
ReDim L2_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L3_LB_000(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L2_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)
ReDim L3_LB_001(Val("&H" & MapLength) - 1, Val("&H" & MapHeight) - 1)

For j = 1 To Val("&H" & MapHeight)
For i = 1 To Val("&H" & MapLength)
L2_LB_000(i - 1, j - 1) = "0000"
L3_LB_000(i - 1, j - 1) = "0000"
L2_LB_001(i - 1, j - 1) = "0000"
L3_LB_001(i - 1, j - 1) = "0000"
Next i
Next j
End If
'end initial

Dim Tilesets As String
Dim StrTemp As String, str1 As String, str2 As String, str3 As String, str4 As String

Form9.Text1.Text = Form9.Text1.Text & "impoting pointers......" & vbCrLf
Tilesets = Mid(Form10.Combo2.Text, 1, 2)
StrTemp = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 9 * 4 + Val("&H" & "3F2298")), Hex(Val("&H" & Tilesets) * 9 * 4 + 35 + Val("&H" & "3F2298")))

TileOffset = Mid(StrTemp, 5, 2) & Mid(StrTemp, 3, 2) & Mid(StrTemp, 1, 2)
TileLength2 = Val("&H" & Mid(StrTemp, 11, 2) & Mid(StrTemp, 9, 2))
Form9.Text1.Text = Form10.Text1.Text & "The amount of Tiles is " & str(TileLength2 / 32) & vbCrLf
TextMAPDataOffset = Mid(StrTemp, 45, 2) & Mid(StrTemp, 43, 2) & Mid(StrTemp, 41, 2)
paletteOffset = Mid(StrTemp, 21, 2) & Mid(StrTemp, 19, 2) & Mid(StrTemp, 17, 2)
ReDim Palette256(16, 16)
ReDim Tile88(TileLength2 / 16 + 64)
Form9.Text1.Text = Form9.Text1.Text & "impoting 8 * 8 Tiles Data......" & vbCrLf
Dim TextMapData As String
StrTemp = ReadFileHexWithByteInterchange(gbafilepath, TileOffset, Hex(Val("&H" & TileOffset) + TileLength2))
'-------------------------------------------------------------权宜之策
str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F8098")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F8098")))
'str1 = ReadFileHex(gbafilepath, Hex(Val("&H" & Tilesets) * 2 ^ 5 + Val("&H" & "3F91D8")), Hex(Val("&H" & Tilesets) * 2 ^ 5 + 31 + Val("&H" & "3F91D8")))
For i = 0 To 15
str2 = Mid(str1, 4 * i + 1, 2)
Mid(str1, 4 * i + 1, 2) = Mid(str1, 4 * i + 3, 2)
Mid(str1, 4 * i + 3, 2) = str2
Next i
For i = 0 To 15
str2 = Mid(str1, 4 * i + 1, 4)
str3 = ReadFileHex(gbafilepath, Hex(Val("&H" & str2) * 8 + Val("&H" & "3F7828")), Hex(Val("&H" & str2) * 8 + Val("&H" & "3F782F")))
str2 = Mid(str3, 9, 8)
str4 = Mid(str3, 1, 2)
str3 = Mid(str3, 5, 2)
If str4 = "03" Or str4 = "06" Then
str3 = Hex(Val("&H" & str3) - 1)
Else
str3 = "00"
End If
str2 = Mid(str2, 7, 2) & Mid(str2, 5, 2) & Mid(str2, 3, 2) & Mid(str2, 1, 2)
str2 = Hex(Val("&H" & str2) - Val("&H8000000") + Val("&H" & str3) * 128)
str2 = ReadFileHexWithByteInterchange(gbafilepath, str2, Hex(Val("&H" & str2) + 127))
DoEvents
For j = 0 To 3
Tile88(i * 4 + j) = Mid(str2, j * 64 + 1, 64)
Next j
Next i
Tile88(64) = Replace(Space(64), Chr(32), "0")
'-------------------------------------------------------------End  权宜之策
For i = 0 To (TileLength2 / 16) - 1
Tile88(i + 65) = Mid(StrTemp, 64 * i + 1, 64)
DoEvents
Next i


Form9.Text1.Text = Form9.Text1.Text & "impoting and making 16 * 16 Tiles Data......" & vbCrLf
TextMapData = ReadFileHex(gbafilepath, TextMAPDataOffset, Hex(Val("&H" & TextMAPDataOffset) + 8192))
For i = 0 To 4095
StrTemp = Mid(TextMapData, 4 * i + 1, 2)
Mid(TextMapData, 4 * i + 1, 2) = Mid(TextMapData, 4 * i + 3, 2)
Mid(TextMapData, 4 * i + 3, 2) = StrTemp
DoEvents
Next i

ReDim Tile16(1024)
Dim r0 As Long, r1 As Long, r2 As Long
For i = 0 To 1023
r0 = i * 4
r2 = r0 Or 1
Tile16(i) = Mid(TextMapData, r0 * 2 * 2 + 1, 4)
r1 = (r2 + 1) * 2 ^ 16
Tile16(i) = Tile16(i) & Mid(TextMapData, r2 * 2 * 2 + 1, 4)
r0 = 128 * 2 ^ 9
r2 = r1 + r0
r1 = RSH(r1, 15)
Tile16(i) = Tile16(i) & Mid(TextMapData, r1 * 2 + 1, 4)
r2 = RSH(r2, 15)
Tile16(i) = Tile16(i) & Mid(TextMapData, r2 * 2 + 1, 4)
DoEvents
Next i

Form9.Text1.Text = Form9.Text1.Text & "impoting palette 256 ......" & vbCrLf
StrTemp = ReadFileHex(gbafilepath, paletteOffset, Hex(Val("&H" & paletteOffset) + 256 * 2 - 1))
For j = 0 To 15
For i = 0 To 15
Palette256(i, j) = RGB555ToRGB888(Mid(StrTemp, 64 * j + 4 * i + 1, 4))
DoEvents
Next i
Next j

If IsDeliver = False Then
Form9.Text1.Text = Form9.Text1.Text & "No Map Text Data inport, create new map!" & vbCrLf
'make grid for Tile16
Form9.Text1.Text = Form9.Text1.Text & "Making grid......" & vbCrLf
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
Form10.Picture1.Line (64 * 6 * i, 64 * 6 * j)-(64 * 6 * i + 64 * 6, 64 * 6 * j + 64 * 6), vbWhite, B
DoEvents
Next i
Next j
'grid end            so each point = 24*24
End If

If IsDeliver = True Then
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
result = DrawTile16(i, j, L1_LB_000(i, j), Form10.Picture1)
DoEvents
Next i
Next j
End If
Form9.Text1.Text = Form9.Text1.Text & "Finish All" & vbCrLf
Form10.Combo1.Enabled = True
End Sub


Private Sub Command6_Click()
Dim i As Integer, j As Integer, result As Boolean
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
L1_LB_000(i, j) = L1_LB_001(i, j)
L1_LB_001(i, j) = "0000"
Next i
Next j
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1)
DoEvents
Next i
Next j
Form10.Command6.Enabled = False
End Sub

Private Sub Command7_Click()
Xshift = Val("&H" & Form10.Text6.Text)
Yshift = Val("&H" & Form10.Text7.Text)
Form10.Picture1.Cls
Dim i As Integer, j As Integer, result As Boolean
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1)
DoEvents
Next i
Next j
End Sub

Private Sub Command8_Click()
Form10.Picture1.Cls
End Sub


Private Sub Command9_Click()
Form10.Picture1.Cls
Form10.Picture1.DrawWidth = 2
Dim i As Integer, j As Integer, result As Boolean
For j = 0 To Val("&H" & MapHeight) - 1 - Yshift
For i = 0 To Val("&H" & MapLength) - 1 - Xshift
result = DrawTile16(i, j, L1_LB_000(i + Xshift, j + Yshift), Form10.Picture1)
DoEvents
Next i
Next j
If Mid(LevelAllRoomPointerandDataallHex, 1 + 48 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2) = "03" Then
    Dim FirstPointer As String
    FirstPointer = Hex(Val("&H" & "78F540") + 4 * Val("&H" & LevelNumber))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 3))
    FirstPointer = Mid(FirstPointer, 7, 2) & Mid(FirstPointer, 5, 2) & Mid(FirstPointer, 3, 2) & Mid(FirstPointer, 1, 2)
    FirstPointer = Hex(Val("&H" & FirstPointer) - Val("&H" & "8000000"))
    FirstPointer = ReadFileHex(gbafilepath, FirstPointer, Hex(Val("&H" & FirstPointer) + 17 * 4 - 1))
    Dim OutputString As String, CheckPointer As String, kk As Integer
    Dim b0 As Integer, b1 As Integer, b2 As Integer, b3 As Integer, b4 As Integer, b5 As Integer
    For i = 0 To 16
    If Mid(FirstPointer, 8 * i + 1, 8) = "589D3F08" Then Exit For
    CheckPointer = Mid(FirstPointer, 7 + 8 * i, 2) & Mid(FirstPointer, 5 + 8 * i, 2) & Mid(FirstPointer, 3 + 8 * i, 2) & Mid(FirstPointer, 1 + 8 * i, 2)
    CheckPointer = Hex(Val("&H" & CheckPointer) - Val("&H" & "8000000"))
    OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 1))
        If Mid(OutputString, 1, 2) = Right("00" & Hex(Val("&H" & LevelRoomIndex) - 1), 2) Then
            RoomCameraStringPointerOffset = CheckPointer
            OutputString = ReadFileHex(gbafilepath, CheckPointer, Hex(Val("&H" & CheckPointer) + 10 * 9 + 1))
            kk = Val("&H" & Mid(OutputString, 3, 2))
            For j = 0 To (kk - 1)
            b0 = Val("&H" & Mid(OutputString, 18 * j + 7, 2))
            b1 = Val("&H" & Mid(OutputString, 18 * j + 9, 2))
            b2 = Val("&H" & Mid(OutputString, 18 * j + 11, 2))
            b3 = Val("&H" & Mid(OutputString, 18 * j + 13, 2))
            Form10.Picture1.Line ((b0 - Xshift) * 24 * 16, (b2 - Yshift) * 24 * 16)-((b1 + 1 - Xshift) * 24 * 16, (b3 + 1 - Yshift) * 24 * 16), vbRed, B
            If Mid(OutputString, 18 * j + 15, 2) <> "FF" Then
            b4 = Val("&H" & Mid(OutputString, 18 * j + 15, 2))
            b5 = Val("&H" & Mid(OutputString, 18 * j + 17, 2))
            Form10.Picture1.Line ((b4 - Xshift) * 24 * 16, (b5 - Yshift) * 24 * 16)-((b4 + 1 - Xshift) * 24 * 16, (b5 + 1 - Yshift) * 24 * 16), vbRed, B
            If Val("&H" & Mid(OutputString, 18 * j + 19, 2)) = "00" Then b0 = Val("&H" & Mid(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid(OutputString, 18 * j + 19, 2)) = "01" Then b1 = Val("&H" & Mid(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid(OutputString, 18 * j + 19, 2)) = "02" Then b2 = Val("&H" & Mid(OutputString, 18 * j + 21, 2))
            If Val("&H" & Mid(OutputString, 18 * j + 19, 2)) = "03" Then b3 = Val("&H" & Mid(OutputString, 18 * j + 21, 2))
            Form10.Picture1.Line ((b0 - Xshift) * 24 * 16, (b2 - Yshift) * 24 * 16)-((b1 + 1 - Xshift) * 24 * 16, (b3 + 1 - Yshift) * 24 * 16), vbRed, B
            End If
            Next j
            Exit For
        End If
    Next i
End If
End Sub

Private Sub Form_Load()
If MODfilepath = "" Then
MsgBox "No MOD file Loaded", vbInformation, "Info"
Form10.Visible = False
Exit Sub
End If

Dim i As Integer, j As Integer
ReDim TileMOD(2, 500)
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
If IsDeliver = True Then
ReDim L1_LB_000(Val("&H" & widtha1) - 1, Val("&H" & heighta2) - 1)
ReDim L1_LB_001(Val("&H" & widtha1) - 1, Val("&H" & heighta2) - 1)
For j = 1 To Val("&H" & heighta2)
For i = 1 To Val("&H" & widtha1)
L1_LB_000(i - 1, j - 1) = "0000"
L1_LB_001(i - 1, j - 1) = "0000"
Next i
Next j
Form10.Text8.Text = widtha1
Form10.Text9.Text = heighta2

For j = 0 To Val("&H" & heighta2) - 1
For i = 0 To Val("&H" & widtha1) - 1
Mid(L1_LB_000(i, j), 3, 2) = Mid(Hexstream1, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2)
Next i
Next j

If Hexstream2 <> "" Then
For j = 0 To Val("&H" & heighta2) - 1
For i = 0 To Val("&H" & widtha1) - 1
 Mid(L1_LB_000(i, j), 1, 2) = Mid(Hexstream2, i * 2 + j * Val("&H" & widtha1) * 2 + 1, 2)
Next i
Next j
End If

Form10.Combo2.ListIndex = Val("&H" & Mid(LevelAllRoomPointerandDataallHex, 1 + (Val("&H" & LevelRoomIndex) - 1) * 44 * 2, 2))

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase Tile16()
Erase Tile88()
Erase Palette256()
'Erase TextMap()

Erase L1_LB_000()
Erase L2_LB_000()
Erase L3_LB_000()
Erase L1_LB_001()
Erase L2_LB_001()
Erase L3_LB_001()

Erase TileMOD()
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X \ (24 * 16)
MouseY = Y \ (24 * 16)
Form10.Shape1.Left = X
Form10.Shape1.Top = Y
End Sub

Private Sub Picture1_Click()
Dim i As Integer, j As Integer
For j = 0 To Val("&H" & MapHeight) - 1
For i = 0 To Val("&H" & MapLength) - 1
L1_LB_001(i, j) = L1_LB_000(i, j)
Next i
Next j
For j = 0 To UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) - 1
For i = 0 To UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) - 1
If NowTileMOD(i, j) <> "0000" Then
result = DrawTile16(MouseX + i, MouseY + j, NowTileMOD(i, j), Form10.Picture1)
L1_LB_000(MouseX + Xshift + i, MouseY + Yshift + j) = NowTileMOD(i, j)
Form10.Command6.Enabled = True
End If
Next i
Next j

If (UBound(NowTileMOD, 1) - LBound(NowTileMOD, 1) = 1) And (UBound(NowTileMOD, 2) - LBound(NowTileMOD, 2) = 1) And NowTileMOD(0, 0) = "0000" Then
result = DrawTile16(MouseX, MouseY, NowTileMOD(0, 0), Form10.Picture1)
L1_LB_001(MouseX + Xshift, MouseY + Yshift) = L1_LB_000(MouseX + Xshift, MouseY + Yshift)
L1_LB_000(MouseX + Xshift, MouseY + Yshift) = NowTileMOD(0, 0)
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
result = DrawTile16(i, j, NowTileMOD(i, j), Form10.Picture2)
Next i
Next j
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X \ (24 * 16)
MouseY = Y \ (24 * 16)
End Sub
