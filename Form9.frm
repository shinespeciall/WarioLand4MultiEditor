VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "output"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   5205
   Begin VB.CommandButton Command1 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Text1.Text = ""
If MODfilepath <> "" Then Form9.Text1.Text = "Load MOD File, now you can make room visually!!" & vbCrLf
End Sub

Private Sub Form_Load()
Form9.Text1.FontSize = 14
Form9.Icon = LoadResPicture(101, vbResIcon)
Form9.Move 0, 9705, 4650, 4540  'Form4.Widht = 4650
If MODfilepath <> "" Then Form9.Text1.Text = "Load MOD File, now you can make room visually!!" & vbCrLf
End Sub

Private Sub Form_Resize()
Form9.Text1.width = Form9.width - 450
End Sub
