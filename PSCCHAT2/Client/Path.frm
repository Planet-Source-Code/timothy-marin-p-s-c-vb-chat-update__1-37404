VERSION 5.00
Begin VB.Form Path 
   Caption         =   "Select Path"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Path.frx":0000
      Top             =   3960
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decline"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2940
      Width           =   4695
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2340
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Select Path To Save File..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public namer, iper, porter As String
Public indexer As Integer
Dim a As Integer
Private Sub Command1_Click(Index As Integer)
pm(indexer).Show
pm(indexer).Caption = "Chat With " & namer
pm(indexer).Winsock2.Connect iper, porter
a = 1
Unload Me
End Sub
Private Sub Command2_Click()
namer = Left(namer, Len(namer) - 1)
namer = Right(namer, Len(namer) - 1)
a = 1
Form1.Winsock1.SendData "PMSG|¿|" & namer & "|¿|" & "/SENDDECLINE"
Unload Me
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Text1.Text = Dir1.Path
Form1.save1 = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
If a = 1 Then GoTo quiter:
Cancel = 1
quiter:
End Sub
