VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "ProFile"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Save Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Password :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Skillz :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Website :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Email :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Age :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "NickName :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = vbChecked Then
    Form1.Passhold.Text = Text1(5).Text
Else
    Form1.Passhold.Text = ""
End If
Form1.Text3.Text = Text1(0).Text
Form1.List2.Clear
For i = 1 To 4
On Error Resume Next
Form1.List2.AddItem Label1(i).Caption & " " & Text1(i).Text
Next
Form1.pass = Text1(5).Text
If Form1.pass > 0 Then
    Form1.Winsock1.Close
    Form1.Winsock1.Connect Form1.Text2.Text ',remote port set in properties
    Form1.s.Panels.Item(1).Text = "Connecting..."
    Form1.kol = 0
    Unload Me
Else
MsgBox "invald password (must be at least 1 chr)"
End If
End Sub

Private Sub Form_Load()
If Len(Form1.Passhold.Text) > 1 Then
    Text1(5).Text = Form1.Passhold.Text
    Check1.Value = vbChecked
End If
Text1(0).Text = Form1.Text3.Text
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

