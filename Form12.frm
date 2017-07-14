VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Model Maker"
   ClientHeight    =   8760
   ClientLeft      =   5760
   ClientTop       =   3660
   ClientWidth     =   10740
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10740
   Begin VB.CommandButton Command4 
      Caption         =   "Next"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Last"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   8160
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   4095
      Left            =   4920
      ScaleHeight     =   4035
      ScaleWidth      =   5595
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   7800
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Save in Block Model file"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6960
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Save in BG Model file"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   4395
      TabIndex        =   0
      ToolTipText     =   "Click and choose"
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Page: 0"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Tiles"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Nmae It !"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TilePage As Integer

Private Sub Command3_Click()
Form12.Picture1.Cls
TilePage = TilePage - 1
Dim a As Boolean, i As Integer, j As Integer
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i, j, Hex(128 * TilePage + i + 8 * j), Form12.Picture1, , 24)
DoEvents
Next i
Next j
If TilePage = 0 Then
Form12.Command3.Enabled = False
Else
Form12.Command3.Enabled = True
End If
Form12.Label3.Caption = "Page: " & TilePage
End Sub

Private Sub Command4_Click()
Form12.Picture1.Cls
TilePage = TilePage + 1
Dim a As Boolean, i As Integer, j As Integer
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i, j, Hex(128 * TilePage + i + 8 * j), Form12.Picture1, , 24)
DoEvents
Next i
Next j
Form12.Command3.Enabled = True
Form12.Label3.Caption = "Page: " & TilePage
End Sub

Private Sub Form_Load()
If TilePage > 0 Then TilePage = TilePage - 1
Form12.Label1.FontSize = 12
Form12.Label2.FontSize = 12
Form12.Label3.FontSize = 12
Form12.Text1.FontSize = 12
Form12.Picture1.BackColor = &H0&
Form12.Command3.Enabled = False

Form12.height = 9345
Form12.width = 10980
Form12.Icon = LoadResPicture(101, vbResIcon)
Form12.Top = MDIForm1.height / 2 - Form12.height / 2 - 500
Form12.Left = MDIForm1.width / 2 - Form12.width / 2 - 1000

Dim a As Boolean, i As Integer, j As Integer
For j = 0 To 15
For i = 0 To 7
a = DrawTile16(i, j, Hex(i + 8 * j), Form12.Picture1, , 24)
DoEvents
Next i
Next j
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form12.Visible = False
Form10.Enabled = True
End Sub
