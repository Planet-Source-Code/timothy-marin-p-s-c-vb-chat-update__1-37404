VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "New User"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Pass"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "User_Name"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "2 Accounts Per Ip !"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Form1.Winsock1.State = 7 Then
Form1.Winsock1.SendData "NEWS|¿|" & Text1.Text & "|¿|" & Text2.Text
End If
End Sub
