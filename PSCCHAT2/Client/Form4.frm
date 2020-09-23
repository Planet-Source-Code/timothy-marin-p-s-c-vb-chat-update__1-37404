VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form4 
   Caption         =   "New Code"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form4"
   ScaleHeight     =   4635
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Navigate Press Enter"
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Code"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5655
      ExtentX         =   9975
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
w.Navigate "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
End Sub

Private Sub Form_Load()
w.Navigate "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Command1.Width = Form4.ScaleWidth - 20
Text1.Width = Form4.ScaleWidth - 20
w.Height = Form4.ScaleHeight - 550
w.Width = Form4.ScaleWidth
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
On Error Resume Next
w.Navigate Text1.Text
End If
End Sub
