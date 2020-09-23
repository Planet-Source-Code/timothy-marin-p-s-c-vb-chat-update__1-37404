VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   Caption         =   "PSC VB CHAT"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Passhold 
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer PingSafety 
      Interval        =   60000
      Left            =   600
      Top             =   480
   End
   Begin VB.ComboBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":08CA
      Left            =   120
      List            =   "Form1.frx":08D4
      TabIndex        =   6
      Text            =   "pscode.no-ip.com"
      Top             =   120
      Width           =   8775
   End
   Begin MSComctlLib.StatusBar s 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15425
            Text            =   "Disconected..."
            TextSave        =   "Disconected..."
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   7080
      TabIndex        =   3
      Text            =   "NickName"
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4695
      Left            =   7080
      TabIndex        =   2
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   8281
      View            =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3088
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   15001
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
      MaxLength       =   200
      TabIndex        =   1
      Top             =   5280
      Width           =   8775
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   7200
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":08FC
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
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   4800
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   2400
      Left            =   7200
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer m 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnucon 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuNews 
         Caption         =   "New User"
      End
      Begin VB.Menu mnudis 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnucoder 
         Caption         =   "View Latest VB Code"
      End
      Begin VB.Menu mnuserch 
         Caption         =   "Code Search"
      End
      Begin VB.Menu mnuFTP 
         Caption         =   "Chat Ftp Server"
      End
      Begin VB.Menu mnuMS 
         Caption         =   "Message Service"
         Begin VB.Menu mnuMSS 
            Caption         =   "Send"
         End
         Begin VB.Menu mnuMSV 
            Caption         =   "View"
         End
      End
      Begin VB.Menu mnuACC 
         Caption         =   "Account"
         Begin VB.Menu mnuPASS 
            Caption         =   "Change Password"
         End
         Begin VB.Menu mnuREMOVE 
            Caption         =   "Remove Account"
         End
      End
      Begin VB.Menu mnuopi 
         Caption         =   "Ops"
         Begin VB.Menu mnuCRB 
            Caption         =   "Clear 60 Min Bans"
         End
         Begin VB.Menu mnuCSB 
            Caption         =   "Clear SubNet Bans"
         End
         Begin VB.Menu mnuCPB 
            Caption         =   "Clear Permanent Bans"
         End
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuVI 
         Caption         =   "View Ignored List"
      End
      Begin VB.Menu mnucI 
         Caption         =   "Clear Ignored"
      End
      Begin VB.Menu mnuEd 
         Caption         =   "Edit Profile"
      End
      Begin VB.Menu mnucolors 
         Caption         =   "Colors"
      End
   End
   Begin VB.Menu mnuMore 
      Caption         =   "&More"
      Begin VB.Menu mnuIPM 
         Caption         =   "Ignore Private Msg's"
      End
      Begin VB.Menu mnuDEC 
         Caption         =   "Decline File Transfers"
      End
      Begin VB.Menu mnuNotify 
         Caption         =   "Notify When Minamized"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbo 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuCommands 
         Caption         =   "Commands"
      End
   End
   Begin VB.Menu mnuA 
      Caption         =   "UserMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuIG 
         Caption         =   "Ignore"
      End
      Begin VB.Menu mnuPM 
         Caption         =   "Private Msg"
      End
      Begin VB.Menu mnuUn 
         Caption         =   "UnIgnore"
      End
      Begin VB.Menu mnuIS 
         Caption         =   "WhoIs"
      End
      Begin VB.Menu mnuOPC 
         Caption         =   "Op Controls"
         Begin VB.Menu mnuKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Ban"
         End
         Begin VB.Menu mnuMBAN 
            Caption         =   "Ban SubNet"
         End
         Begin VB.Menu mnuVIP 
            Caption         =   "View Ip"
         End
         Begin VB.Menu mnuMuzzle 
            Caption         =   "Muzzle"
         End
         Begin VB.Menu mnuBanN 
            Caption         =   "Ban Name"
         End
      End
   End
   Begin VB.Menu mnuj 
      Caption         =   "j"
      Visible         =   0   'False
      Begin VB.Menu mnuSHOW 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuHIDE 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuCLS 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public kol As Integer
Public save1 As String
Private Const WM_PASTE = &H302
Public pass As String
Const Similys = ":),;),:z,:d,:(,:|,:m,:<,:o,:x"
Public mecolor, maintext As ColorConstants
Dim sss(50) As New Path
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Dim WithEvents SysIcon As CSystrayIcon  'Create an instance of CSystrayIcon using events
Attribute SysIcon.VB_VarHelpID = -1

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'When the callback message of CSystrayIcon is WM_MOUSEMOVE,
  'the X of Form_MouseMove is used to see what happen to the
  'icon in the systray.
  Dim msgCallBackMessage As Long
  
  'To be able to compare the callback value to the window message,
  'we must divide X by Screen.TwipsPerPixelX. That represent the
  'horizontal number of twips in the screen. (1 pixel ~= 15 twips)
  msgCallBackMessage = x / Screen.TwipsPerPixelX
   
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
    WindowState = vbNormal
    Me.Show
    Case WM_LBUTTONUP
    
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN

    Case WM_RBUTTONUP
    PopupMenu mnuj
    Case WM_RBUTTONDBLCLK
    Case WM_MBUTTONDOWN
    Case WM_MBUTTONUP
    Case WM_MBUTTONDBLCLK
  End Select
End Sub
Private Sub mnuCLS_Click()
On Error Resume Next
SysIcon.HideIcon
Open App.Path & "\Colors.ini" For Output As #122
Print #122, Form1.Passhold.Text
Print #122, Form1.Text3.Text
Print #122, Form1.mecolor
Print #122, Form1.maintext
Print #122, Form1.LV1.BackColor
Print #122, Form1.BackColor
Print #122, Form1.LV1.ForeColor
Print #122, Form1.Rtb.BackColor
Print #122, Form1.Text1.BackColor
Print #122, Form1.Text1.ForeColor
Print #122, Form1.Text2.BackColor
Print #122, Form1.Text2.ForeColor
For pop = 0 To List2.ListCount - 1
On Error Resume Next
Print #122, List2.List(pop)
Next
Close #122
Dim lngWin32apiResultCode As Long

lngWin32apiResultCode = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, glnglpOriginalWndProc)
End
End Sub

Private Sub mnuHide_Click()
WindowState = vbMinimized
End Sub

Private Sub mnuShow_Click()
    WindowState = vbNormal
    Me.Show
End Sub

Private Sub SysIcon_NIError(ByVal ErrorNumber As Long)
  MsgBox "There was an error when trying to use the systray. #ERR=" & ErrorNumber
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WindowState = vbMinimized Then
    Me.Hide
    Exit Sub
End If
If Form1.Width < 5000 Or Form1.Height < 3000 Then
Form1.Width = 5000
Form1.Height = 3000
End If
Text2.Width = Form1.ScaleWidth - 210
Rtb.Width = Text2.Width - (LV1.Width + 150)
Text3.Left = Form1.ScaleWidth - (Text3.Width + 100)
LV1.Left = Text3.Left
Text1.Width = Rtb.Width + LV1.Width + 120
Text1.Top = Form1.ScaleHeight - 650
LV1.Height = Form1.ScaleHeight - (Text1.Height + 950)
Rtb.Height = LV1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.WindowState = vbMinimized
End Sub
Private Sub Form_Load()
Form1.save1 = App.Path
Dim lngEventMask   As Long
Dim lngWin32apiResultCode As Long
  Set SysIcon = New CSystrayIcon 'Set the new instance
  SysIcon.Initialize hWnd, Me.Icon, "VB Chat"
  SysIcon.ShowIcon
With Rtb
    
    
    lngEventMask = SendMessage(.hWnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
    
    If lngEventMask Xor ENM_LINK Then
        lngEventMask = lngEventMask Or ENM_LINK
    End If
    lngWin32apiResultCode = SendMessage(.hWnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
    lngWin32apiResultCode = SendMessage(.hWnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
    
End With
mecolor = &H80FF&
mintext = vbBlack
On Error GoTo hyhy
Open App.Path & "\Colors.ini" For Input As #342
Line Input #342, col
Form1.Passhold.Text = col
Line Input #342, col
Form1.Text3.Text = col
Line Input #342, col
Form1.mecolor = col
Line Input #342, col
Form1.maintext = col
Line Input #342, col
Form1.LV1.BackColor = col
Line Input #342, col
Form1.BackColor = col
Line Input #342, col
Form1.LV1.ForeColor = col
Line Input #342, col
Form1.Rtb.BackColor = col
Line Input #342, col
Form1.Text1.BackColor = col
Line Input #342, col
Form1.Text1.ForeColor = col
Line Input #342, col
Text3.BackColor = col
Text2.BackColor = col
Line Input #342, col
Form1.Text2.ForeColor = col
List2.Clear
Line Input #342, col
List2.AddItem col
Line Input #342, col
List2.AddItem col
Line Input #342, col
List2.AddItem col
Line Input #342, col
List2.AddItem col
hyhy:
Close #342

glngOriginalhWnd = Me.hWnd

glnglpOriginalWndProc = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, AddressOf RichTextBoxSubProc)

End Sub

Private Sub LV1_DblClick()
mnuPM_Click
End Sub

Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuA
End If
End Sub

Private Sub mnuAbo_Click()
Form2.Show
End Sub
Private Sub mnuBanN_Click()
If Winsock1.State = 7 Then
    If Len(LV1.SelectedItem) > 0 Then
        Winsock1.SendData "OPS|¿|NAME|¿|" & LV1.SelectedItem
    End If
End If
End Sub
Private Sub mnuBan_Click()
If Winsock1.State = 7 Then
    If Len(LV1.SelectedItem) > 0 Then
        Winsock1.SendData "OPS|¿|BAN|¿|" & LV1.SelectedItem
    End If
End If
End Sub

Private Sub mnucI_Click()
List1.Clear
End Sub

Private Sub mnucoder_Click()
Form4.Show
End Sub

Private Sub mnucolors_Click()
Form6.Show
End Sub

Private Sub mnuCommands_Click()
MsgBox "/clear - Clear you chat screen..." & vbCrLf & vbCrLf & "!Seen Name - see when a user last logged in..." & vbCrLf & vbCrLf & "Emoticons - :),;),:z,:d,:(,:|,:m,:<,:o,:x" & vbCrLf & vbCrLf & "/me - First person msg..." & vbCrLf & vbCrLf & "/Ops - See Online Ops..." & vbCrLf & vbCrLf, , "Commands"
End Sub

Private Sub mnucon_Click()
'check to see if connected already
'if not connect
If Winsock1.State <> 7 Then
Form7.Show
End If
End Sub

Private Sub mnuCPB_Click()
If Winsock1.State = 7 Then
    Winsock1.SendData "OPS|¿|CPB"
End If
End Sub

Private Sub mnuCRB_Click()
If Winsock1.State = 7 Then
    Winsock1.SendData "OPS|¿|CRB"
End If
End Sub

Private Sub mnuCSB_Click()
If Winsock1.State = 7 Then
    Winsock1.SendData "OPS|¿|CSB"
End If
End Sub

Private Sub mnuDEC_Click()
    If mnuDEC.Checked = True Then
        mnuDEC.Checked = False
    Else
        mnuDEC.Checked = True
    End If
End Sub

Private Sub mnudis_Click()
'close the winsock
'let user know what happend
Winsock1.Close
Rtb.Text = "Disconnected..."
s.Panels.Item(1).Text = "Disconnected..."
LV1.ListItems.Clear
End Sub

Private Sub mnuEd_Click()
Form8.Show
End Sub

Private Sub mnuFTP_Click()
Form4.Show
Form4.w.Navigate "ftp://psc@pscode.no-ip.com"
End Sub

Private Sub mnuIG_Click()
On Error Resume Next
List1.AddItem "<" & LV1.SelectedItem & ">"
End Sub

Private Sub mnuIPM_Click()
    If mnuIPM.Checked = True Then
        mnuIPM.Checked = False
    Else
        mnuIPM.Checked = True
    End If
End Sub

Private Sub mnuIS_Click()
If Winsock1.State = 7 Then
Winsock1.SendData "PMSG|¿|" & LV1.SelectedItem & "|¿|/WHOAMI"
List3.AddItem "<" & LV1.SelectedItem & ">"
End If
End Sub

Private Sub mnuKick_Click()
If Winsock1.State = 7 Then
    If Len(LV1.SelectedItem) > 0 Then
        Winsock1.SendData "OPS|¿|KICK|¿|" & LV1.SelectedItem
    End If
End If
End Sub
Private Sub mnuMBAN_Click()
If Winsock1.State = 7 Then
    If Len(LV1.SelectedItem) > 0 Then
        Winsock1.SendData "OPS|¿|SUBNET|¿|" & LV1.SelectedItem
    End If
End If
End Sub

Private Sub mnuMSS_Click()
    Dim namemsg As String
    Dim msgmsg As String
    namemsg = InputBox("Please Type the users name (must be valid).", "UserName")
    If Len(namemsg) < 3 Then Exit Sub
    msgmsg = InputBox("Please Type the msg to send.", "")
    If Len(msgmsg) < 3 Then Exit Sub
    If Winsock1.State = 7 Then
        Winsock1.SendData "OMSG|¿|SEND|¿|" & namemsg & "|¿|" & msgmsg
    End If
End Sub

Private Sub mnuMSV_Click()
If Winsock1.State = 7 Then
        Winsock1.SendData "OMSG|¿|LIST|¿|hmmok"
End If
End Sub

Private Sub mnuMuzzle_Click()
If Winsock1.State = 7 Then
Dim mtime As String
mtime = InputBox("Tpye Muzzle Length In Seconds...", "Time", "60")
If Len(mtime) < 1 Then Exit Sub
If IsNumeric(mtime) = True Then
    If Len(LV1.SelectedItem.Text) > 0 Then
        Winsock1.SendData "OPS|¿|MUZZLE|¿|" & LV1.SelectedItem.Text & "|¿|" & mtime
    End If
End If
End If
End Sub

Private Sub mnuNews_Click()
On Error Resume Next
    Form1.Winsock1.Close
    Form1.Winsock1.Connect Form1.Text2.Text ',remote port set in properties
    Form1.s.Panels.Item(1).Text = "Connecting..."
    kol = 1
End Sub

Private Sub mnuNotify_Click()
    If mnuNotify.Checked = True Then
        mnuNotify.Checked = False
    Else
        mnuNotify.Checked = True
    End If
End Sub

Private Sub mnuPASS_Click()
Dim passer As String
Dim checker As String
passer = InputBox("Select a new password...", "Change Pass")
If Len(passer) < 3 Then GoTo erro
checker = InputBox("Select a new password...", "Change Pass")
If checker = passer Then
If Winsock1.State = 7 Then
    Winsock1.SendData "ACCOUNT|¿|PASS|¿|" & passer
    Exit Sub
End If
End If
erro:
MsgBox "There was an error changing your password.", , "Error"
End Sub

Private Sub mnuPM_Click()
If LV1.ListItems.Count = 0 Then Exit Sub
For i = 1 To 50
If pm(i).Visible = True Then
If InStr(pm(i).Caption, "<" & LV1.SelectedItem & ">") Then
pm(i).Show
Exit Sub
End If
Else
pm(i).Show
pm(i).Caption = "Chat With " & "<" & LV1.SelectedItem & ">"
Exit Sub
End If
Next
'Form3.Show
'Form3.Combo1.Clear
'For i = 1 To Form1.LV1.ListItems.Count
'    Form3.Combo1.AddItem Form1.LV1.ListItems.Item(i)
'Next
'Form3.Combo1.ListIndex = LV1.SelectedItem.Index - 1
End Sub

Private Sub mnuRemove_Click()
Dim passer As String
passer = InputBox("Type yes to continue...", "Remove Account")
If LCase(passer) = "yes" Then
If Winsock1.State = 7 Then
    Winsock1.SendData "ACCOUNT|¿|REMOVE"
    Exit Sub
End If
End If
erro:
MsgBox "There was an error removing your account.", , "Error"
End Sub

Private Sub mnuserch_Click()
Dim a As String
a = InputBox("Search For...", "Search", "Timothy Marin")
If Len(a) > 0 Then
Form4.Show
Form4.w.Navigate "http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=1&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & a
End If
End Sub

Private Sub mnuUn_Click()
For i = 0 To List1.ListCount - 1
    If List1.List(i) = "<" & LV1.SelectedItem & ">" Then
        List1.RemoveItem (i)
    End If
Next
End Sub

Private Sub mnuvI_Click()
Dim a As String
For i = 0 To List1.ListCount - 1
    a = a & List1.List(i) & vbCrLf
Next
MsgBox a, , "Ignored List..."
End Sub

Private Sub mnuVIP_Click()
If Winsock1.State = 7 Then
    If Len(LV1.SelectedItem) > 0 Then
        Winsock1.SendData "OPS|¿|VIEW|¿|" & LV1.SelectedItem
    End If
End If
End Sub

Private Sub PingSafety_Timer()
If Winsock1.State = 7 Then
Winsock1.SendData "PING|¿|" & Time
End If
End Sub

Private Sub rtb_Change()
On Error Resume Next
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
    If LCase(Left(Text1.Text, 6)) = "/clear" Then
        Text1.Text = ""
        Rtb.Text = ""
        Exit Sub
    End If
    If UCase(Text1.Text) = "GBL" Or UCase(Text1.Text) = "GSBL" Or UCase(Text1.Text) = "GPBL" Then
        If Winsock1.State = 7 Then
            Winsock1.SendData "OPS|¿|" & UCase(Text1.Text) & vbCrLf
        End If
        Text1.Text = ""
    End If
    If Winsock1.State = 7 Then
        Winsock1.SendData "MSG|¿|" & Text1.Text
    End If
    Text1.Text = ""
End If
End Sub



Private Sub Text3_Change()
'dont remove this - u will be banned
Text3.Text = Replace(Text3.Text, " ", "")
Text3.Text = Replace(Text3.Text, ">", "")
Text3.Text = Replace(Text3.Text, "<", "")
Text3.Text = Replace(Text3.Text, "Þ", "")
End Sub




Private Sub Winsock1_Close()
'close the winsock
'let user know what happend
Winsock1.Close
Rtb.Text = "Disconnected..."
s.Panels.Item(1).Text = "Disconnected..."
LV1.ListItems.Clear
End Sub

Private Sub Winsock1_Connect()
'tell user they connected
Rtb.Text = "Connected..."
s.Panels.Item(1).Text = "Connected..."
'send server nickname and wait for list
If kol = 1 Then
Form9.Show
kol = 0
Exit Sub
End If

Winsock1.SendData "LOGIN|¿|" & Text3.Text & "|¿|" & pass
LV1.ListItems.Clear
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

' dim var for parsing data
Dim split1, split2, split3 As Variant
' dim a string and get the data sent
Dim sata As String
Winsock1.GetData sata
'Rtb.Text = Rtb.Text & vbCrLf & sata
split1 = Split(sata, vbCrLf)
For x = 0 To UBound(split1)
On Error Resume Next
split3 = Split(split1(x), "|¿|")
Select Case split3(0)
    Case "OMSG"
    If UBound(split3) < 4 Then GoTo Shooter22
    frmOffline.Show
    frmOffline.MS1.ListItems.Add , , split3(1)
    frmOffline.MS1.ListItems.Item(frmOffline.MS1.ListItems.Count).SubItems(1) = split3(2)
    frmOffline.MS1.ListItems.Item(frmOffline.MS1.ListItems.Count).SubItems(2) = split3(3)
    frmOffline.MS1.ListItems.Item(frmOffline.MS1.ListItems.Count).SubItems(3) = split3(4)
    GoTo Shooter22
    Case "PING"
        Winsock1.SendData "PING|¿|" & Time & vbCrLf
    Case "INUSE"
        MsgBox "Unable To Create New Account... Possible Resons, Invalid Name, Name in Use, You Own 2 Accounts... Please Check this and Try Again"
    Case "DONER"
        MsgBox "New Acount Created Please Go To File/Connect And Login"
        Unload Form9
        Winsock1_Close
        Exit Sub
    Case "INVALID"
        MsgBox "Unable To Login Please Check Settings And Try Again"
        Winsock1_Close
        Exit Sub
    Case "LIST"
        LV1.ListItems.Add , , split3(1)
    Case "JOIN"
        Rtb.SelStart = Len(Rtb.Text)
        LV1.ListItems.Add , , split3(1)
        Rtb.SelColor = &H8000&
        Rtb.SelBold = True
        Rtb.SelText = vbCrLf & " * " & split3(1) & " Has Joined..."
        'Winsock1.SendData "PMSG|¿|" & split3(1) & "|¿|" & "Do To Varius Bugs For The moment i have switched back to an older version... get it here : ftp://psc@pscode.no-ip.com look for new_code.zip"
        If mnuNotify.Checked = True Then
            On Error Resume Next
            If Me.WindowState = vbMinimized Then

                m.FileName = App.Path & "\gling.wav"
                m.Play
            End If

        End If

    Case "PART"
    Rtb.SelStart = Len(Rtb.Text)
        For k = 1 To LV1.ListItems.Count
            If LV1.ListItems.Item(k) = split3(1) Then
                LV1.ListItems.Remove k
                Rtb.SelColor = vbRed
                Rtb.SelBold = True
                Rtb.SelText = vbCrLf & " * " & split3(1) & " Has Quit..."
                Exit Sub
            End If
        Next
        
    Case "NICK"
    Dim newn As String
here:
    newn = InputBox("Please Select a New Nick...", "NewNick", "NewNick")
    If Len(newn) < 1 Then GoTo here
    Text3.Text = newn
    Winsock1.SendData "NICK|¿|" & Text3.Text
    Exit Sub
    Case "MSG"
        
    Rtb.SelStart = Len(Rtb.Text)
    Dim aka As Variant
    aka = Split(split3(1), " ")
    For i = 0 To List1.ListCount - 1
        If aka(0) = List1.List(i) Then Exit Sub
    Next
    If mnuNotify.Checked = True Then
    On Error Resume Next
        If Me.WindowState = vbMinimized Then
            m.FileName = App.Path & "\gling.wav"
            m.Play
        End If
    End If
        
    If LCase(aka(1)) = "/me" Then
        Dim det As String
        For hjk = 2 To UBound(aka)
            det = det & " " & aka(hjk)
        Next
        Rtb.SelColor = mecolor
        Rtb.SelBold = True
        aka(0) = Replace(aka(0), "<", "")
        aka(0) = Replace(aka(0), ">", "")
        Rtb.SelText = vbCrLf & " * " & aka(0) & det
        Exit Sub
    End If
    Rtb.SelColor = maintext
    Rtb.SelText = vbCrLf & split3(1)
    
    Case "SMSG"
    MsgBox split3(1), , "SERVER MSG"
    
    Case "PMSG"
    Dim aka1 As Variant
    aka1 = Split(split3(1), " ")
    For g = 0 To List1.ListCount - 1
        If aka1(0) = List1.List(g) Then
            aka1(0) = Replace(aka1(0), "<", "")
            aka1(0) = Replace(aka1(0), ">", "")
            Winsock1.SendData "PMSG|¿|" & aka1(0) & "|¿|I am ignoring you..."
            Exit Sub
        End If
    Next
        If mnuIPM.Checked = True Then
            aka1(0) = Replace(aka1(0), "<", "")
            aka1(0) = Replace(aka1(0), ">", "")
            Winsock1.SendData "PMSG|¿|" & aka1(0) & "|¿|I am ignoring all private msg's..."
            Exit Sub
        End If
    On Error Resume Next
        For hg = 1 To 50
            If pm(hg).Visible = True Then
                If InStr(pm(hg).Caption, aka1(0)) Then
                        Dim kok, kok1 As Variant
                        kok = Split(split3(1), " ")
                        If Left(kok(1), 9) = "/SENDFILE" Then
send:
                            For i = 3 To UBound(kok)
                                On Error Resume Next
                                kok(2) = kok(2) & kok(i)
                            Next
                            kok1 = Split(kok(2), "|@|")
                            Dim ad As Integer
                            kok(0) = Left(kok(0), Len(kok(0)) - 1)
                            kok(0) = Right(kok(0), Len(kok(0)) - 1)
                            If mnuDEC.Checked = True Then
                                Form1.Winsock1.SendData "PMSG|¿|" & kok(0) & "|¿|" & "/SENDDECLINE"
                                Exit Sub
                            End If
                       '     ad = MsgBox(kok(0) & " Would Like To Send You " & kok1(0) & " Will You Accept?", vbYesNo, "FILE...")
                       '     If ad = 6 Then
                       '         pm(hg).Show
                       '         pm(hg).Caption = "Chat With " & aka1(0)
                       '         pm(hg).Winsock2.Connect kok1(2), kok1(1)
                       '         Exit Sub
                       '     End If
                       '     Form1.Winsock1.SendData "PMSG|¿|" & kok(0) & "|¿|" & "/SENDDECLINE"
                       '     Exit Sub
                       ' End If
                       Dim kloj As Integer
                       For kloj = 0 To 50
                       If Len(kok1(0)) < 5 Then
                        Winsock1.SendData "PMSG|¿|" & kok(0) & "|¿|" & "/SENDDECLINE"
                        Exit Sub
                       End If
                            If sss(kloj).Visible = False Then
                                sss(kloj).Show
                                sss(kloj).Text2.Text = "Info : " & kok(0) & " Would Like To Send You " & kok1(0) & " Will You Accept?"
                                sss(kloj).indexer = hg
                                sss(kloj).namer = aka1(0)
                                sss(kloj).iper = kok1(2)
                                sss(kloj).porter = kok1(1)
                                Exit Sub
                            End If
                        Next
                        Form1.Winsock1.SendData "PMSG|¿|" & kok(0) & "|¿|" & "/SENDDECLINE"
                    End If
                    If Left(kok(1), 12) = "/SENDDECLINE" Then
dec:
                        MsgBox "Transfer Declined...", , "Declined"
                        Call pm(hg).Winsock1_Close
                        Exit Sub
                    End If

                    If Left(kok(1), 7) = "/WHOYOU" Then
WHO:
                    Dim fifa As String
                    fifa = Right(split3(1), Len(split3(1)) - (7 + 2 + Len(kok(0))))
                    kok1 = Split(fifa, "|@|")
                    'rtb.SelStart = Len(rtb.Text)
                    'rtb.SelColor = vbBlue
                    'rtb.SelText = vbCrLf & "WHOIS - " & kok(0)
                    Dim kok22 As String
                    Dim gdf As Integer
                    For gdf = 0 To List3.ListCount - 1
                        If List3.List(gdf) = kok(0) Then
                        List3.RemoveItem gdf
                                For koli = 0 To UBound(kok1)
                                On Error Resume Next
                                'rtb.SelStart = Len(rtb.Text)
                                'rtb.SelColor = vbBlue
                                'rtb.SelText = vbCrLf & kok1(koli)
                                kok22 = kok22 & kok1(koli) & vbCrLf
                            Next
                            MsgBox kok22, , kok(0)
                            Exit Sub
                        End If
                    Next
                    End If
                    If Left(kok(1), 7) = "/WHOAMI" Then
IAM:
                    Dim kinfo As String
                    For kinho = 0 To List2.ListCount - 1
                    On Error Resume Next
                    kinfo = kinfo & List2.List(kinho) & "|@|"
                    Next
                    kok(0) = Left(kok(0), Len(kok(0)) - 1)
                    kok(0) = Right(kok(0), Len(kok(0)) - 1)
                    Winsock1.SendData "PMSG|¿|" & kok(0) & "|¿|/WHOYOU " & kinfo
                    Exit Sub
                    End If
                    pm(hg).Rtb.Text = pm(hg).Rtb.Text & vbCrLf & split3(1)
                    Exit Sub
                Else
                    GoTo here22
                End If
            Else
            kok = Split(split3(1), " ")
            If Left(kok(1), 9) = "/SENDFILE" Then
            GoTo send
            ElseIf Left(kok(1), 12) = "/SENDDECLINE" Then
            GoTo dec:
            ElseIf Left(kok(1), 7) = "/WHOYOU" Then
            GoTo WHO:
            ElseIf Left(kok(1), 7) = "/WHOAMI" Then
            GoTo IAM:
            Else
                                        pm(hg).Show
                pm(hg).Caption = "Chat With " & aka1(0)
                pm(hg).Rtb.Text = pm(hg).Rtb.Text & vbCrLf & split3(1)
                Exit Sub
            End If
            End If
here22:
        Next
    End Select
Shooter22:
Next

'Code For SMileys :)
    Dim NumberOfIcon As Integer
    Dim StartOfIcon As Integer
    tmp = Split(Similys, ",")
    For i = 0 To UBound(tmp)
doover:
        If InStr(LCase(Rtb.Text), LCase(tmp(i))) Then
            DoEvents
            On Error Resume Next
            NumberOfIcon = InStr(LCase(Rtb.Text), LCase(tmp(i)))
            StartOfIcon = NumberOfIcon + 1
            Rtb.SelStart = NumberOfIcon - 1
            Rtb.SelLength = 2
            temp = Clipboard.GetText
            Clipboard.Clear
            Clipboard.SetData LoadResPicture(1 + i, vbResBitmap)
            Rtb.Locked = False
            SendMessage Rtb.hWnd, WM_PASTE, 0, 0
            Rtb.Locked = True
            Clipboard.SetText temp
        End If
        If InStr(1, LCase(Rtb.Text), LCase(tmp(i))) Then GoTo doover:
    Next

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1.Close
    Rtb.Text = Description
    s.Panels.Item(1).Text = "Disconnected..."
    LV1.ListItems.Clear
End Sub

