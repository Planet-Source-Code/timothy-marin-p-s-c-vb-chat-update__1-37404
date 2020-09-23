VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Scheme"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Reset Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   4215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Address/Nick BG"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Address/Nick Text "
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Send Text Color"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Send Text BG"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Text BackGround"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nick Text"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Main BackGround"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nick BackGround"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Main Text Color"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Emotion Color"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   1200
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1680
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1680
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1680
      Top             =   600
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1680
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1680
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Top             =   600
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
C.ShowColor
Shape1.BackColor = C.Color
Form1.mecolor = C.Color
End Sub

Private Sub Command10_Click()
C.ShowColor
Shape9.BackColor = C.Color
Form1.Text2.BackColor = C.Color
Form1.Text3.BackColor = C.Color
End Sub

Private Sub Command11_Click()
Form1.mecolor = 33023
Form1.maintext = 0
Form1.LV1.BackColor = -2147483643
Form1.BackColor = -2147483633
Form1.LV1.ForeColor = -2147483640
Form1.rtb.BackColor = -2147483643
Form1.Text1.BackColor = -2147483643
Form1.Text1.ForeColor = -2147483640
Form1.Text2.BackColor = -2147483643
Form1.Text3.BackColor = -2147483643
Form1.Text2.ForeColor = -2147483640
Form_Load
End Sub

Private Sub Command2_Click()
C.ShowColor
Shape2.BackColor = C.Color
Form1.maintext = C.Color
End Sub

Private Sub Command3_Click()
C.ShowColor
Shape3.BackColor = C.Color
Form1.LV1.BackColor = C.Color
End Sub

Private Sub Command4_Click()
C.ShowColor
Shape4.BackColor = C.Color
Form1.BackColor = C.Color
End Sub

Private Sub Command5_Click()
C.ShowColor
Shape5.BackColor = C.Color
Form1.LV1.ForeColor = C.Color
End Sub

Private Sub Command6_Click()
C.ShowColor
Shape6.BackColor = C.Color
Form1.rtb.BackColor = C.Color
End Sub

Private Sub Command7_Click()
C.ShowColor
Shape7.BackColor = C.Color
Form1.Text1.BackColor = C.Color
End Sub

Private Sub Command8_Click()
C.ShowColor
Shape8.BackColor = C.Color
Form1.Text1.ForeColor = C.Color
End Sub

Private Sub Command9_Click()
C.ShowColor
Shape10.BackColor = C.Color
Form1.Text2.ForeColor = C.Color
Form1.Text3.ForeColor = C.Color
End Sub

Private Sub Form_Load()
On Error Resume Next
Shape1.BackColor = Form1.mecolor
Shape2.BackColor = Form1.maintext
Shape3.BackColor = Form1.LV1.BackColor
Shape4.BackColor = Form1.BackColor
Shape5.BackColor = Form1.LV1.ForeColor
Shape6.BackColor = Form1.rtb.BackColor
Shape7.BackColor = Form1.Text1.BackColor
Shape8.BackColor = Form1.Text1.ForeColor
Shape9.BackColor = Form1.Text2.BackColor
Shape10.BackColor = Form1.Text2.ForeColor
End Sub
