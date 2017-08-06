VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "output"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   5205
   Begin VB.TextBox Text1 
      Height          =   3010
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Double Click to Clear up"
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form9.Text1.FontSize = 14
Form9.Icon = LoadResPicture(101, vbResIcon)
Form9.Move 0, 8950, 4650, 3010
If MODfilepath <> "" Then Form9.Text1.Text = "Load MOD File, now you can make room visually!!" & vbCrLf
End Sub

Private Sub Form_Resize()
Form9.Text1.Width = Form9.Width - 450
Form9.Text1.Height = Form9.Height - 800
End Sub

Private Sub Text1_DblClick()
Form9.Text1.Text = ""
End Sub
