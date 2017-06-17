VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "timer change"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   4335
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "ReSave"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
    For i = 1 To 100
        If SaveDataOffset(i) = "" Then Exit For
    Next i

SaveDataOffset(i) = LevelStartStreamOffset
SaveDatabuffer(i) = Form8.Text1.Text
End Sub

Private Sub Form_Activate()
If LevelNumber = "" Then
Unload Form8
Exit Sub
End If
Form8.Move 4650, 3000, 4650, 4800
Form8.Label1.Caption = "the level number is " & LevelNumber & Chr(13) & "the order is Hard, Normal, S-Hard" & Chr(13) & "Do not change the first three Byte(s) only if you know what are they!"
Form8.Text1.FontSize = 15
Form8.Label1.FontSize = 15

Form8.Text1.Text = LevelStartStream
End Sub

Private Sub Form_Load()
Form8.Visible = False
End Sub

