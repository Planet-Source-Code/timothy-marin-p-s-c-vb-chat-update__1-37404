VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   3960
      Top             =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Form1.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Image1_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Label1_Click()
Shell "explorer.exe http://www.intradream.com"
End Sub

Private Sub Timer1_Timer()
    Unload Me
    Form1.Show
End Sub
