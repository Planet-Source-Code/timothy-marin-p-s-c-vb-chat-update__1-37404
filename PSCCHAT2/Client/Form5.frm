VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   Caption         =   "Send Code"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7080
   LinkTopic       =   "Form5"
   ScaleHeight     =   5895
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info"
      Height          =   2055
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   600
         Top             =   240
      End
      Begin MSComctlLib.ProgressBar p2 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   5
      End
      Begin MSComctlLib.ProgressBar p1 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sending..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.TextBox rtb 
      Height          =   5535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   7095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuSend 
         Caption         =   "&Send"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time1 As Integer
Dim b As Integer


Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
rtb.SetFocus
End Sub

Private Sub Form_Load()
time1 = 0
For i = 1 To Form1.LV1.ListItems.Count
Combo1.AddItem Form1.LV1.ListItems.Item(i)
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtb.Height = Form5.ScaleHeight - 380
rtb.Width = Me.ScaleWidth
Combo1.Width = Me.ScaleWidth - 20
End Sub

Private Sub mnuSend_Click()
If Timer1.Enabled = True Then Exit Sub
If b = 1 Then Exit Sub
b = 1
Dim a As Variant
Frame1.Visible = True
a = Split(rtb.Text, vbCrLf)
On Error Resume Next
p1.Max = UBound(a)
For i = 0 To UBound(a)
On Error Resume Next
p1.Value = i
Form1.Winsock1.SendData "PMSG|¿|" & Combo1.Text & "|¿|" & a(i) & vbCrLf
For lol = 0 To 80000
DoEvents
Next
Next
Timer1.Enabled = True
b = 0
End Sub

Private Sub Timer1_Timer()
time1 = time1 + 1
If time1 > 5 Then
    p1.Value = Min
    Timer1.Enabled = False
    time1 = 0
    Frame1.Visible = False
    Exit Sub
End If
p2.Value = time1
End Sub
