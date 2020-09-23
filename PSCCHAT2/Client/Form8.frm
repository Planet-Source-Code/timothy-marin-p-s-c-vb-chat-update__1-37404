VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Edit Profile"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Age :"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Email :"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Website :"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Skillz :"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.List2.Clear
For i = 1 To 4
On Error Resume Next
Form1.List2.AddItem Label1(i).Caption & " " & Text1(i).Text
Next
Unload Me
End Sub

Private Sub Form_Load()
Dim kik As Variant
For i = 0 To Form1.List2.ListCount - 1
kik = Split(Form1.List2.List(i), " : ")
Dim ojo As String
For g = 1 To UBound(kik)
On Error Resume Next
ojo = ojo & kik(g)
Next
DoEvents
On Error Resume Next
Dim x As Integer
x = i + 1
Text1(x).Text = ojo
ojo = ""
Next
End Sub
