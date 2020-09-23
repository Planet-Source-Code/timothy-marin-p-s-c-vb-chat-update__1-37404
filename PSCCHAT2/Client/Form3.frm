VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Chat With  "
   ClientHeight    =   4200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6690
   LinkTopic       =   "Form3"
   ScaleHeight     =   4200
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Uploading"
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin MSComctlLib.ProgressBar p2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Downlaoding"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4575
      End
      Begin MSComctlLib.ProgressBar p1 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   480
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6376
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   6495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSND 
         Caption         =   "&Send Code Snipit"
      End
      Begin VB.Menu mnuSNDFile 
         Caption         =   "Send File"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Transfers"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Transfers"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim File As String
Public aj As Integer
Dim sends As Integer

Private Sub Command1_Click()
Winsock2_Close
Winsock2.Close
End Sub

Private Sub Command2_Click()
Winsock1_Close
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
Winsock2.Close
End Sub

Private Sub mnuHide_Click()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub mnuShow_Click()
Frame1.Visible = True
Frame2.Visible = True
End Sub

Private Sub Text1_Change()
Text1.Text = Replace(Text1.Text, "/WHOYOU", "/whoyou")
End Sub

Public Sub Winsock1_Close()
Winsock1.Close
aj = 0
Close #1
sends = 0
Frame2.Visible = False
Label2.Caption = "0"
Label5.Caption = "0"
Text2.Text = ""
On Error Resume Next
p2.Value = 0
Label7.Caption = "Disonnected..."
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Frame2.Visible = True
Text3.Text = File
End Sub
Sub sendfile()
Open File For Binary As #1
Dim Data As String
Data = Space$(1028)
sends = 1
Do Until EOF(1)
Do Until sends = 1
If Winsock1.State <> 7 Then
Winsock1_Close
Winsock1.Close
Exit Sub
End If
DoEvents
Loop
If Winsock1.State <> 7 Then
Winsock1_Close
Winsock1.Close
Exit Sub
End If
Label2.Caption = LOF(1)
p2.Max = LOF(1)
Label5.Caption = Label5.Caption + 1028
On Error Resume Next
p2.Value = Label5.Caption
Label7.Caption = "Connected..."
DoEvents
sends = 0
Get #1, , Data
Winsock1.SendData Data
DoEvents
Loop
MsgBox "File Transfer Complete - " & File, , "Complete..."
On Error Resume Next
Winsock1.SendData "DONE"
Frame2.Visible = False
aj = 0
Close #1
sends = 0
Frame2.Visible = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sata As String
Winsock1.GetData sata
If sata = "MORE" Then
sends = 1
End If
If sata = "INFO" Then
Winsock1.SendData "INFO|@|" & File & "|@|" & FileLen(File)
End If
If sata = "SEND" Then
Call sendfile
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
aj = 0
Frame2.Visible = False
Close #1
sends = 0
Label2.Caption = "0"
Label5.Caption = "0"
Text2.Text = ""
On Error Resume Next
p2.Value = 0
Label7.Caption = "Disonnected..."
End Sub
Private Sub Form_Resize()
On Error Resume Next
Randomize
Text1.Top = Me.ScaleHeight - 405
Text1.Width = Me.ScaleWidth - 230
Rtb.Width = Me.ScaleWidth - 230
Rtb.Height = Me.ScaleHeight - (Text1.Height + 380)
End Sub

Private Sub mnuSND_Click()
Form5.Show
    Dim klop As Variant
    klop = Split(Me.Caption, " ")
    klop(UBound(klop)) = Left(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
    klop(UBound(klop)) = Right(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
    Form5.Combo1.Text = klop(UBound(klop))
End Sub

Private Sub mnuSNDFile_Click()
If aj = 1 Then Exit Sub
aj = 1
CommonDialog1.DialogTitle = "Select a File."
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "All Files"
   CommonDialog1.ShowOpen     ' = 1
If CommonDialog1.FileName = "" Then
aj = 0
Exit Sub
End If
File = CommonDialog1.FileName
Dim x As String
here:
Dim MyValue
MyValue = Int((5000 * Rnd) + 200)
x = InputBox("Select a port... Must allow connections through firewall/router...", "PORT", MyValue)
If Len(x) > 0 Then
    If x < 10 Or x > 50000 Then
        MsgBox "Unusable Port"
        GoTo here:
    End If
On Error GoTo error
Winsock1.LocalPort = x
Winsock1.Listen
    If Form1.Winsock1.State = 7 Then
    Dim klop As Variant
    klop = Split(Me.Caption, " ")
    klop(UBound(klop)) = Left(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
    klop(UBound(klop)) = Right(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
        Form1.Winsock1.SendData "PMSG|¿|" & klop(UBound(klop)) & "|¿|/SENDFILE|@|" & File & "|@|" & x
        Rtb.Text = Rtb.Text & vbCrLf & "<" & Form1.Text3.Text & "> " & "Trying to send " & File

    End If
Else
aj = 0
End If
Exit Sub
error:
MsgBox error
aj = 0
End Sub

Private Sub rtb_Change()
Rtb.SelStart = Len(Rtb.Text)
End Sub

Private Sub Rtb_KeyPress(KeyAscii As Integer)
If KeyAscii = 3 Then Exit Sub
Text1.SetFocus
Text1.Text = Text1.Text & Chr(KeyAscii)
Text1.SelStart = Len(Text1.Text)
KeyAscii = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Text1.Text = "" Then Exit Sub
    If LCase(Text1.Text) = "/cancel" Then
    Winsock1_Close
    Winsock1.Close
    Winsock2_Close
    Winsock2.Close
    Text1.Text = ""
    Exit Sub
    End If
    If Form1.Winsock1.State = 7 Then
    Dim klop As Variant
    klop = Split(Me.Caption, " ")
    klop(UBound(klop)) = Left(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
    klop(UBound(klop)) = Right(klop(UBound(klop)), Len(klop(UBound(klop))) - 1)
        Form1.Winsock1.SendData "PMSG|¿|" & klop(UBound(klop)) & "|¿|" & Text1.Text
        Rtb.Text = Rtb.Text & vbCrLf & "<" & Form1.Text3.Text & "> " & Text1.Text
    End If
    Text1.Text = ""
End If
End Sub

Private Sub Winsock2_Close()
Winsock2.Close
Close #2
If Label9.Caption > Label10.Caption - 1 Then
MsgBox "File Transfer Complete", , "Complete..."
Else
MsgBox "File Transfer - Posible Error - Check " & Text2.Text, , "Transfer..."
End If
Frame1.Visible = False
Label9.Caption = "0"
Label10 = "0"
p1.Value = 0
Text2.Text = ""
Label8.Caption = "Disconnected..."
End Sub

Private Sub Winsock2_Connect()
On Error Resume Next
Winsock2.SendData "INFO"
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim sata As String
Winsock2.GetData sata
Dim kook As Variant
kook = Split(sata, "|@|")
If kook(0) = "INFO" Then
On Error Resume Next
Dim name As Variant
name = Split(kook(1), "\")
Kill Form1.save1 & "\" & name(UBound(name))
Open Form1.save1 & "\" & name(UBound(name)) For Binary As #2
Frame1.Visible = True
Label8.Caption = "Connected..."
Text2.Text = Form1.save1 & "\" & name(UBound(name))
Label10.Caption = kook(2)
p1.Max = kook(2)
Winsock2.SendData "SEND"
ElseIf kook(0) = "DONE" Then
Winsock2_Close
Winsock2.Close
Else
Put #2, , sata
Winsock2.SendData "MORE"
On Error GoTo error
DoEvents
Label9.Caption = Label9.Caption + bytesTotal
If Label9.Caption > p1.Max Then
p1.Value = p1.Max
Else
p1.Value = Label9.Caption
End If
'Winsock1.SendData "MORE"
End If
Exit Sub
error:
MsgBox error
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2.Close
Close #2
Frame2.Visible = False
If Label9.Caption > Label10.Caption - 1 Then
MsgBox "File Transfer Complete", , "Complete..."
Else
MsgBox "File Transfer - Posible Error - Check " & Text2.Text, , "Transfer..."
End If
Frame1.Visible = False
Label9.Caption = "0"
p1.Value = 0
Text2.Text = ""
Label8.Caption = "disconnected..."
End Sub
