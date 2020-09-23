VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC SERVER"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "RTB"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LV1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Winsock1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Timer4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flood Catch"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "List6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Timer2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Timer1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "List3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "List2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "List1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Banned"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command8"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "List7"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "List5"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "List4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Command3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Command2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text3"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label3"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label2"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Users"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "U1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "U2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label5"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Muzzle"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Timer5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command10"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "M1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Other"
      TabPicture(5)   =   "Form1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame2"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame Frame4 
         Caption         =   "More Flood"
         Height          =   855
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   6615
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   2520
            TabIndex        =   45
            Text            =   "5"
            Top             =   510
            Width           =   735
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2520
            TabIndex        =   43
            Text            =   "6"
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Kicks Till Ban :"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   540
            Width           =   2295
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Winsock Connections Till Ban :"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Set Bog"
         Height          =   375
         Left            =   6000
         TabIndex        =   41
         Top             =   5340
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4920
         TabIndex        =   40
         Text            =   "4096"
         Top             =   5325
         Width           =   960
      End
      Begin VB.Frame Frame3 
         Caption         =   "Msg Service"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   38
         Top             =   3360
         Width           =   6615
         Begin MSComctlLib.ListView MS1 
            Height          =   2415
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4260
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "To"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "From"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Time/Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Message"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ping Results"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   36
         Top             =   1800
         Width           =   6615
         Begin VB.Timer Timer6 
            Interval        =   60000
            Left            =   6120
            Top             =   240
         End
         Begin MSComctlLib.ListView P1 
            Height          =   1095
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   1931
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Winsock"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Word Filter"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   6660
         Begin VB.ComboBox List8 
            Height          =   315
            ItemData        =   "Form1.frx":00A8
            Left            =   120
            List            =   "Form1.frx":00CA
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   240
            Width           =   6375
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Remove"
            Height          =   255
            Left            =   3360
            TabIndex        =   34
            Top             =   720
            Width           =   3135
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Add"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   3135
         End
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   -69120
         Top             =   2640
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Clear Muzzle's"
         Height          =   315
         Left            =   -74880
         TabIndex        =   31
         Top             =   6000
         Width           =   6615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear Subnets"
         Height          =   375
         Left            =   -69600
         TabIndex        =   27
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72000
         TabIndex        =   26
         Text            =   "127.0"
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Ban"
         Height          =   375
         Left            =   -73440
         TabIndex        =   25
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Unban"
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ListBox List7 
         Height          =   840
         Left            =   -74880
         TabIndex        =   22
         Top             =   2520
         Width           =   6615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear All Bans"
         Height          =   375
         Left            =   -69600
         TabIndex        =   19
         Top             =   5835
         Width           =   1335
      End
      Begin MSComctlLib.ListView U1 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   17
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Login"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pass"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ip"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Last"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Timer Timer4 
         Interval        =   30000
         Left            =   5880
         Top             =   0
      End
      Begin VB.Timer Timer3 
         Interval        =   45000
         Left            =   -73080
         Top             =   3960
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   -74880
         TabIndex        =   16
         Top             =   3195
         Width           =   4095
      End
      Begin VB.Timer Timer2 
         Interval        =   30000
         Left            =   -73560
         Top             =   3960
      End
      Begin VB.ListBox List5 
         Height          =   1425
         Left            =   -74880
         TabIndex        =   15
         Top             =   4200
         Width           =   6615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "x"
         Height          =   255
         Left            =   -68400
         TabIndex        =   14
         Top             =   5940
         Width           =   135
      End
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   -74040
         Top             =   3960
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear All"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   5940
         Width           =   6495
      End
      Begin VB.ListBox List4 
         Height          =   1620
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   6615
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   -74880
         TabIndex        =   11
         Top             =   4080
         Width           =   6615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Unban"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   5820
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ban"
         Height          =   375
         Left            =   -73440
         TabIndex        =   9
         Top             =   5820
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72000
         TabIndex        =   8
         Text            =   "127.0.0.1"
         Top             =   5820
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   -70680
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Text            =   "Topic : Visual Basic"
         Top             =   5340
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Msg"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   5820
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   5820
         Width           =   5175
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   0
         Left            =   6360
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   15001
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   4815
         Left            =   4920
         TabIndex        =   4
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   8493
         View            =   3
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   4815
         Left            =   135
         TabIndex        =   5
         Top             =   405
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8493
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Form1.frx":010E
      End
      Begin MSComctlLib.ListView U2 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   18
         Top             =   4080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Index"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ip"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView M1 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   30
         Top             =   1320
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Online Users"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Accounts"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Banned Subnets"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Permanent Bans"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   3960
         Width           =   6615
      End
      Begin VB.Label Label1 
         Caption         =   "Bans Cleared Eery 60 Min."
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PoP"
      Visible         =   0   'False
      Begin VB.Menu mnuRsel 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuRALL 
         Caption         =   "Remove All"
      End
   End
   Begin VB.Menu m 
      Caption         =   "m"
      Visible         =   0   'False
      Begin VB.Menu mnubsnet 
         Caption         =   "Ban Subnet"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban User"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick User"
      End
      Begin VB.Menu mnuWarn 
         Caption         =   "Warn User"
      End
   End
   Begin VB.Menu n 
      Caption         =   "n"
      Visible         =   0   'False
      Begin VB.Menu mnuRemacc 
         Caption         =   "Remove Account"
      End
      Begin VB.Menu mnuBanip 
         Caption         =   "Ban Ip"
      End
      Begin VB.Menu mnubsnet1 
         Caption         =   "Ban Subnet"
      End
      Begin VB.Menu mnuOpUser 
         Caption         =   "Op User"
      End
      Begin VB.Menu mnuDeop 
         Caption         =   "DeOp User"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nick() As String
Dim poop As Integer
Dim PingTime As Integer
Dim lngsock As Long
Dim flooder1 As Integer
Dim dater As String
Dim bog() As Long
Dim setbog As Long
'Window messages that identify mouse action
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

Private Sub Command10_Click()
M1.ListItems.Clear
End Sub

Private Sub Command11_Click()
Dim add1 As String
add1 = InputBox("What To Filter...", "Filter", "Swear")
If add1 > "" Then
    List8.AddItem add1
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
List8.RemoveItem List8.ListIndex
End Sub

Private Sub Command13_Click()
On Error Resume Next
If IsNumeric(Text5.Text) = True Then
    setbog = Text5.Text
Else
Text5.Text = setbog
End If
End Sub

Private Sub Command6_Click()
    List4.Clear
    List5.Clear
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    List5.RemoveItem List7.ListIndex
End Sub

Private Sub Command8_Click()
    List7.AddItem Text4.Text
End Sub

Private Sub Command9_Click()
    List7.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim msgCallBackMessage As Long
  msgCallBackMessage = x / Screen.TwipsPerPixelX
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
        WindowState = vbNormal
        Me.Show
    Case WM_LBUTTONUP
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
        WindowState = vbNormal
        Me.Show
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
    Case WM_MBUTTONDOWN
    Case WM_MBUTTONUP
    Case WM_MBUTTONDBLCLK
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Cancel = 1
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub mnuBanip_Click()
    On Error Resume Next
        List5.AddItem U1.ListItems.Item(U1.SelectedItem.Index).SubItems(3)
End Sub

Private Sub mnuBsnet_Click()
    On Error Resume Next
    Dim subn As Variant
    Dim subn1 As String
    subn = Split(U2.ListItems.Item(U2.SelectedItem.Index).SubItems(3), ".")
    subn1 = subn(0) & "." & subn(1)
        List7.AddItem subn1
        Winsock1_Close (U2.ListItems.Item(U2.SelectedItem.Index).SubItems(2))
End Sub

Private Sub mnubsnet1_Click()
    On Error Resume Next
    Dim subn As Variant
    Dim subn1 As String
    subn = Split(U1.ListItems.Item(U1.SelectedItem.Index).SubItems(3), ".")
    subn1 = subn(0) & "." & subn(1)
            List7.AddItem subn1
End Sub

Private Sub mnuDeop_Click()
    U1.SelectedItem.SubItems(2) = "0"
End Sub

Private Sub mnuOpUser_Click()
    U1.SelectedItem.SubItems(2) = "1"
End Sub

Private Sub mnuRALL_Click()
MS1.ListItems.Clear
End Sub

Private Sub mnuRemacc_Click()
    On Error Resume Next
    U1.ListItems.Remove U1.SelectedItem.Index
End Sub



Private Sub mnuRsel_Click()
On Error Resume Next
MS1.ListItems.Remove MS1.SelectedItem.Index
End Sub

Private Sub MS1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuPop
End If
End Sub

Private Sub SysIcon_NIError(ByVal ErrorNumber As Long)
    MsgBox "There was an error when trying to use the systray. #ERR=" & ErrorNumber
End Sub

Private Sub Command1_Click()
    Call sendall("SMSG|¿|" & Text2.Text)
End Sub

Private Sub Command2_Click()
List5.AddItem Text3.Text
End Sub

Private Sub Command3_Click()
On Error Resume Next
List5.RemoveItem List5.ListIndex
End Sub

Private Sub Command4_Click()
List3.Clear
List4.Clear
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim newdate
If dater = Date Then
newdate = Replace(Date, "/", ".")
Open App.Path & "\" & newdate & ".txt" For Output As #1
Print #1, Time
Print #1, RTB.Text
Close #1
Else
dater = Date
RTB.Text = ""
End If
Open App.Path & "\Users\Accounts.ini" For Output As #1
For i = 1 To U1.ListItems.Count
Print #1, U1.ListItems.Item(i) & "|¿|" & U1.ListItems.Item(i).SubItems(1) & "|¿|" & U1.ListItems.Item(i).SubItems(2) & "|¿|" & U1.ListItems.Item(i).SubItems(3) & "|¿|" & U1.ListItems.Item(i).SubItems(4) & "|¿|" & U1.ListItems.Item(i).SubItems(5)
Next
Close #1
SysIcon.HideIcon
End
End Sub

Private Sub Form_Load()
setbog = 4096
dater = Date
poop = 0
Winsock1(0).Listen
  Set SysIcon = New CSystrayIcon 'Set the new instance
  SysIcon.Initialize hWnd, Me.Icon, "VB Chat Server."
  SysIcon.ShowIcon
  Load Winsock1(1)
  'On Error Resume Next
 Open App.Path & "\Users\Accounts.ini" For Input As #1
Dim im As String
Dim pute As Variant
Do Until EOF(1)
DoEvents
Line Input #1, im
pute = Split(im, "|¿|")
If UBound(pute) > 4 Then
U1.ListItems.Add , , pute(0)
U1.ListItems.Item(U1.ListItems.Count).SubItems(1) = pute(1)
U1.ListItems.Item(U1.ListItems.Count).SubItems(2) = pute(2)
U1.ListItems.Item(U1.ListItems.Count).SubItems(3) = pute(3)
U1.ListItems.Item(U1.ListItems.Count).SubItems(4) = pute(4)
U1.ListItems.Item(U1.ListItems.Count).SubItems(5) = pute(5)
End If
Loop
Close #1
End Sub

Private Sub mnuBan_Click()
On Error Resume Next
        List5.AddItem U2.ListItems.Item(U2.SelectedItem.Index).SubItems(3)
        Winsock1_Close (U2.ListItems.Item(U2.SelectedItem.Index).SubItems(2))
End Sub

Private Sub mnuKick_Click()
On Error Resume Next
        Winsock1_Close (U2.ListItems.Item(U2.SelectedItem.Index).SubItems(2))
End Sub

Private Sub mnuWarn_Click()
On Error Resume Next
Dim warn As String
warn = InputBox("Type Warning To Send...", "Warning", "You have been warned.. one more time you will be kicked/banned.")
If Len(warn) > 0 Then
Winsock1(U2.ListItems.Item(U2.SelectedItem.Index).SubItems(2)).SendData "SMSG|¿|" & warn
End If
End Sub

Private Sub RTB_Change()
RTB.SelStart = Len(RTB.Text)
End Sub




Private Sub Text6_Change()
If IsNumeric(Text6.Text) = False Then
    Text6.Text = "6"
End If
End Sub

Private Sub Text7_Change()
If IsNumeric(Text6.Text) = False Then
    Text6.Text = "5"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
poop = poop + 1
If poop > 179 Then
On Error Resume Next
List3.Clear
List4.Clear
poop = 0
End If
End Sub

Private Sub Timer2_Timer()
List2.Clear
End Sub

Private Sub Timer3_Timer()
List6.Clear
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
Dim newdate
If RTB.Text = "" Then Exit Sub
If dater = Date Then
newdate = Replace(Date, "/", ".")
Open App.Path & "\" & newdate & ".txt" For Append As #1
Print #1, RTB.Text
Close #1
RTB.Text = ""
Else
dater = Date
RTB.Text = ""
End If
Open App.Path & "\Users\Accounts.ini" For Output As #1
For i = 1 To U1.ListItems.Count
Print #1, U1.ListItems.Item(i) & "|¿|" & U1.ListItems.Item(i).SubItems(1) & "|¿|" & U1.ListItems.Item(i).SubItems(2) & "|¿|" & U1.ListItems.Item(i).SubItems(3) & "|¿|" & U1.ListItems.Item(i).SubItems(4) & "|¿|" & U1.ListItems.Item(i).SubItems(5)
Next
Close #1
For i = 1 To UBound(bog)
    On Error Resume Next
    bog(i) = 0
Next
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
    For i = 1 To M1.ListItems.Count
        M1.ListItems.Item(i).SubItems(1) = M1.ListItems.Item(i).SubItems(1) - 1
        If M1.ListItems.Item(i).SubItems(1) < 1 Then
            M1.ListItems.Remove i
        End If
    Next
    On Error Resume Next
    For i = 1 To P1.ListItems.Count
        P1.ListItems.Item(i).SubItems(1) = P1.ListItems.Item(i).SubItems(1) - 1
        If P1.ListItems.Item(i).SubItems(1) < 1 Then
            On Error Resume Next
            Winsock1_Close Int(P1.ListItems.Item(i))
            P1.ListItems.Remove i
        End If
    Next
End Sub

Private Sub Timer6_Timer()
PingTime = PingTime + 1
If PingTime > 1 Then
    On Error Resume Next
    For i = 1 To U2.ListItems.Count
        On Error Resume Next
        If Winsock1(U2.ListItems.Item(i).SubItems(2)).State = 7 Then
            On Error Resume Next
            Winsock1(U2.ListItems.Item(i).SubItems(2)).SendData "PING|¿|" & Time & vbCrLf
            P1.ListItems.Add , , U2.ListItems.Item(i).SubItems(2)
            P1.ListItems.Item(i).SubItems(1) = "90"
        End If
    Next
    PingTime = 0
End If
End Sub



Private Sub U1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
U1.Sorted = True
U1.Sorted = False
End Sub

Private Sub U1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu n
End If
End Sub

Private Sub U2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu m
End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
On Error Resume Next
For i = 1 To P1.ListItems.Count
On Error Resume Next
    If P1.ListItems.Item(i) = Index Then
        P1.ListItems.Remove i
    End If
Next
For i = 1 To LV1.ListItems.Count
On Error Resume Next
    If LV1.ListItems.Item(i) = nick(Index) Then
    On Error Resume Next
    LV1.ListItems.Remove i
    End If
Next
For i = 1 To U2.ListItems.Count
On Error Resume Next
    If U2.ListItems.Item(i) = nick(Index) Then
    Dim vg
    For vg = 1 To U1.ListItems.Count
        If U1.ListItems.Item(vg) = nick(Index) Then
        U1.ListItems.Item(vg).SubItems(4) = Time
        U1.ListItems.Item(vg).SubItems(5) = Date
        End If
    Next
    On Error Resume Next
    U2.ListItems.Remove i
    End If
Next
done:
bog(Index) = 0
If nick(Index) = "" Then
On Error Resume Next
Winsock1(Index).Close
Exit Sub
End If
On Error Resume Next
RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & Winsock1(Index).RemoteHostIP & " Disconnected...(" & Index & ")"
RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & nick(Index) & " Left...(" & Index & ")"


Call sendall("PART|¿|" & nick(Index))
nick(Index) = ""
Winsock1(Index).Close
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If flooder1 = 1 Then
        Exit Sub
    End If
    
    flooder1 = 1
    Dim socket, x, y As Long
    Dim xi As Integer
    Dim k As Integer
    k = 0
    
    For x = 2 To Winsock1.UBound
        If Winsock1(x).State <> 7 Then
            On Error Resume Next
            Winsock1(x).Close
            Winsock1(x).Accept requestID
            On Error Resume Next
            If List6.List(0) = Winsock1(x).RemoteHostIP Then
                List6.AddItem Winsock1(x).RemoteHostIP
                If List6.ListCount > Text6.Text - 1 Then
                    List4.AddItem Winsock1(x).RemoteHostIP
                    Call Winsock1_Close(Int(x))
                    List6.Clear
                End If
            Else
                List6.Clear
                List6.AddItem Winsock1(x).RemoteHostIP
            End If
            RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & Winsock1(x).RemoteHostIP & " : Connectected... (" & x & ")"
            ReDim Preserve nick(Winsock1.UBound)
            ReDim Preserve bog(Winsock1.UBound)
            For y = 1 To Winsock1.UBound
            If Winsock1(y).State = 7 Then
            On Error Resume Next
                If Winsock1(y).RemoteHostIP = Winsock1(x).RemoteHostIP Then
                On Error Resume Next
                    k = k + 1
                        If k > 2 Then
                        On Error Resume Next
                            Call Winsock1_Close(Int(x))
                            Winsock1(x).Close
                        End If
                    End If
                End If
            Next
            For y = 0 To List4.ListCount - 1
                On Error Resume Next
                If List4.List(y) = Winsock1(x).RemoteHostIP Then
                    On Error Resume Next
                    Call Winsock1_Close(Int(x))
                    Winsock1(x).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            For y = 0 To List5.ListCount - 1
            On Error Resume Next
                If List5.List(y) = Winsock1(x).RemoteHostIP Then
                    Call Winsock1_Close(Int(x))
                    Winsock1(x).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            For y = 0 To List7.ListCount - 1
            On Error Resume Next
                If Left(Winsock1(x).RemoteHostIP, Len(List7.List(y))) = List7.List(y) Then
                On Error Resume Next
                    Call Winsock1_Close(Int(x))
                    Winsock1(x).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            Dim kl As Integer
            kl = tally(Winsock1(x).RemoteHostIP)
            If kl > 0 Then
                Winsock1(x).SendData "SMSG|¿|WARNING " & kl & Text7.Text & " OF " & " KICKS TILL BAN.." & vbCrLf
            End If
            flooder1 = 0
            Exit Sub
        End If
    Next
    On Error Resume Next
    socket = Winsock1.UBound + 1
    On Error Resume Next
    Load Winsock1(socket)
    Winsock1(socket).Close
    Winsock1(socket).Accept requestID
            On Error Resume Next
            If List6.List(0) = Winsock1(socket).RemoteHostIP Then
                List6.AddItem Winsock1(socket).RemoteHostIP
                If List6.ListCount > Text6.Text - 1 Then
                    List4.AddItem Winsock1(socket).RemoteHostIP
                    Call Winsock1_Close(Int(socket))
                    List6.Clear
                End If
            Else
            List6.Clear
            List6.AddItem Winsock1(socket).RemoteHostIP
            End If
    ReDim Preserve nick(Winsock1.UBound)
    ReDim Preserve bog(Winsock1.UBound)
    RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & Winsock1(socket).RemoteHostIP & " : Connectected...(" & socket & ")"
    For y = 1 To Winsock1.UBound
    On Error Resume Next
        If Winsock1(y).State = 7 Then
        On Error Resume Next
            If Winsock1(y).RemoteHostIP = Winsock1(socket).RemoteHostIP Then
                k = k + 1
                If k > 2 Then
                    Call Winsock1_Close(Int(socket))
                    Winsock1(socket).Close
                End If
            End If
        End If
    Next
            For y = 0 To List4.ListCount - 1
            On Error Resume Next
                If List4.List(y) = Winsock1(socket).RemoteHostIP Then
                    Call Winsock1_Close(Int(socket))
                    Winsock1(socket).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            On Error Resume Next
            For y = 0 To List5.ListCount - 1
            On Error Resume Next
                If List5.List(y) = Winsock1(socket).RemoteHostIP Then
                    On Error Resume Next
                    Call Winsock1_Close(Int(socket))
                    Winsock1(socket).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            For y = 0 To List7.ListCount - 1
            On Error Resume Next
                If Left(Winsock1(socket).RemoteHostIP, Len(List7.List(y))) = List7.List(y) Then
                    On Error Resume Next
                    Call Winsock1_Close(Int(socket))
                    Winsock1(socket).Close
                    flooder1 = 0
                    Exit Sub
                End If
            Next
            On Error Resume Next
            kl = tally(Winsock1(socket).RemoteHostIP)
            If kl > 0 Then
                On Error Resume Next
                Winsock1(socket).SendData "SMSG|¿|WARNING " & kl & " OF 5 KICKS TILL BAN.." & vbCrLf
            End If
    flooder1 = 0
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim split1 As Variant
    Dim sata As String
    On Error Resume Next
    Winsock1(Index).GetData sata
    bog(Index) = bog(Index) + bytesTotal
    'RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] BOG : " & Winsock1(Index).RemoteHostIP & " : " & nick(Index) & " : " & Index & " : " & bog(Index)
    If bog(Index) > setbog Then
    List3.AddItem Winsock1(Index).RemoteHostIP
                                    MS1.ListItems.Add , , nick(Index)
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(1) = "SERVERÞ"
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(2) = Date & " : " & Time
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(3) = "YOU WHERE KICKED : SERVER BOG > " & setbog & "BYTES IN 30 SECONDS!"
        Winsock1_Close Index
        Exit Sub
    End If
    sata = Replace(sata, "Â", "") ' for flash xmlsocket client's
    Dim liner As Variant
    Dim jo
    liner = Split(sata, vbCrLf)
    On Error Resume Next
    For koper = 0 To UBound(liner)
        If Len(liner(koper)) < 1 Then GoTo Youknow:
            split1 = Split(liner(koper), "|¿|")
            Select Case split1(0)
            Case "PING"
            For jo = 1 To P1.ListItems.Count
            On Error Resume Next
             If P1.ListItems.Item(jo) = Index Then
                P1.ListItems.Remove jo
             End If
            Next
            Case "OMSG"
                If UBound(split1) < 1 Then Exit Sub
                Select Case split1(1)
                    Case "SEND"
                    If UBound(split1) < 3 Then Exit Sub
                        For re = 1 To U1.ListItems.Count
                            If LCase(split1(2)) = LCase(U1.ListItems.Item(re)) Then
                                MS1.ListItems.Add , , split1(2)
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(1) = nick(Index)
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(2) = Date & " : " & Time
                                MS1.ListItems.Item(MS1.ListItems.Count).SubItems(3) = split1(3)
                                Winsock1(Index).SendData "SMSG|¿|Msg Sent..."
                                Exit Sub
                            End If
                        Next
                        Winsock1(Index).SendData "SMSG|¿|Error Sending Your Msg... User Not Registered."
                    Case "REMOVE"
                    If UBound(split1) < 2 Then Exit Sub
                    On Error Resume Next
                        If LCase(MS1.ListItems.Item(Int(split1(2)))) = LCase(nick(Index)) Then
                            MS1.ListItems.Remove Int(split1(2))
                            Winsock1(Index).SendData "SMSG|¿|Msg Removed..." & vbCrLf
                        GoTo chitter
                        End If
                            Winsock1(Index).SendData "SMSG|¿|Error Removing Msg..." & vbCrLf
chitter:
                    Case "LIST"
                        For re = 1 To MS1.ListItems.Count
                            If LCase(nick(Index)) = LCase(MS1.ListItems.Item(re)) Then
                                Winsock1(Index).SendData "OMSG|¿|" & re & "|¿|" & MS1.ListItems.Item(re).SubItems(1) & "|¿|" & MS1.ListItems.Item(re).SubItems(2) & "|¿|" & MS1.ListItems.Item(re).SubItems(3) & vbCrLf
                            End If
                        Next
                End Select
            Case "ACCOUNT"
            If UBound(split1) < 1 Then Exit Sub
                Select Case split1(1)
                    Case "REMOVE"
                     Dim zxc As Integer
                        For zxc = 1 To U1.ListItems.Count
                        On Error Resume Next
                           If LCase(U1.ListItems.Item(zxc)) = LCase(nick(Index)) Then
                               U1.ListItems.Remove zxc
                               Winsock1(Index).SendData "SMSG|¿|Account " & nick(Index) & " Removed Succesfully."
                               Exit Sub
                            End If
                        Next
                        If UBound(split1) < 2 Then GoTo Youknow
                    Case "PASS"
                   
                        For zxc = 1 To U1.ListItems.Count
                         On Error Resume Next
                           If LCase(U1.ListItems.Item(zxc)) = LCase(nick(Index)) Then
                           If Len(split1(2)) < 3 Then GoTo kkkl
                               U1.ListItems.Item(zxc).SubItems(1) = split1(2)
                               Winsock1(Index).SendData "SMSG|¿|Account " & nick(Index) & " New Password Set : " & split1(2)
                               GoTo Youknow
                        End If
                        Next
                End Select
kkkl:
Winsock1(Index).SendData "SMSG|¿|Error Editing Account."
            Case "NEWS"
            If UBound(split1) < 2 Then Exit Sub
            On Error Resume Next
            split1(1) = Replace(split1(1), " ", "")
            For re = 0 To 39
                DoEvents
                split1(1) = Replace(split1(1), Chr(re), "")
                split1(2) = Replace(split1(2), Chr(re), "")
            Next
            For re = 123 To 255
                DoEvents
                split1(1) = Replace(split1(1), Chr(re), "")
                split1(2) = Replace(split1(2), Chr(re), "")
            Next
                If Len(split1(1)) < 3 Or Len(split1(2)) < 2 Then
                    Winsock1(Index).SendData "INUSE|¿|654" & vbCrLf
                    Exit Sub
                End If
                If Len(split1(1)) > 15 Or Len(split1(2)) > 15 Then
                    Winsock1(Index).SendData "INUSE|¿|654" & vbCrLf
                    Exit Sub
                End If
                
                For jo = 1 To U1.ListItems.Count
                    If U1.ListItems.Item(jo).SubItems(3) = Winsock1(Index).RemoteHostIP Then
                        agw = agw + 1
                        If agw > 1 Then
                            Winsock1(Index).SendData "INUSE|¿|670" & vbCrLf
                            Exit Sub
                        End If
                    End If
                Next
            For jo = 1 To U1.ListItems.Count
                If UCase(split1(1)) = UCase(U1.ListItems.Item(jo)) Then
                    Winsock1(Index).SendData "INUSE|¿|669" & vbCrLf
                    Exit Sub
                End If
            Next
            agw = 0
            U1.ListItems.Add , , split1(1)
            U1.ListItems.Item(U1.ListItems.Count).SubItems(1) = split1(2)
            U1.ListItems.Item(U1.ListItems.Count).SubItems(2) = "0"
            U1.ListItems.Item(U1.ListItems.Count).SubItems(3) = Winsock1(Index).RemoteHostIP
            U1.ListItems.Item(U1.ListItems.Count).SubItems(4) = Time
            U1.ListItems.Item(U1.ListItems.Count).SubItems(5) = Date
            Winsock1(Index).SendData "DONER|¿|669" & vbCrLf
        
            Case "LOGIN"
            For jo = 1 To U2.ListItems.Count
                If U2.ListItems.Item(jo).SubItems(2) = Index Then
                    Winsock1_Close Index
                    Exit Sub
                End If
            Next
                If UBound(split1) < 1 Then Exit Sub
                For jo = 1 To U2.ListItems.Count
                    If UCase(split1(1)) = UCase(U2.ListItems.Item(jo)) Then
                        Winsock1(Index).SendData "INVALID|¿|669" & vbCrLf
                        Exit Sub
                    End If
                Next
                For jo = 1 To U1.ListItems.Count
                On Error Resume Next
                    If UCase(split1(1)) = UCase(U1.ListItems.Item(jo)) Then
                    If UBound(split1) < 2 Then Exit Sub
                    Dim banname As Integer
                    For banname = 0 To List4.ListCount - 1
                        If UCase(split1(1)) = UCase(List4.List(banname)) Then
                            Winsock1(Index).Close
                            Exit Sub
                        End If
                    Next
                        If split1(2) = U1.ListItems.Item(jo).SubItems(1) Then
                        Winsock1(Index).SendData "SMSG|¿|" & Text1.Text & vbCrLf
                                U2.ListItems.Add , , split1(1)
                                U2.ListItems.Item(U2.ListItems.Count).SubItems(1) = U1.ListItems.Item(jo).SubItems(2)
                                U2.ListItems.Item(U2.ListItems.Count).SubItems(2) = Index
                                U2.ListItems.Item(U2.ListItems.Count).SubItems(3) = Winsock1(Index).RemoteHostIP
                                U1.ListItems.Item(jo).SubItems(4) = Time
                                U1.ListItems.Item(jo).SubItems(5) = Date
                                nick(Index) = split1(1)
                                For i = 1 To LV1.ListItems.Count
                                    Winsock1(Index).SendData "LIST|¿|" & LV1.ListItems.Item(i) & vbCrLf
                                Next
                                LV1.ListItems.Add , , nick(Index)
                                Dim Numder As Integer
                                Numder = 0
                                On Error Resume Next
                                For re = 1 To MS1.ListItems.Count
                                    If LCase(MS1.ListItems.Item(re)) = LCase(nick(Index)) Then
                                     Numder = Numder + 1
                                    End If
                                Next
                                 Winsock1(Index).SendData "SMSG|¿|You Have " & Numder & " Messages [ Tools | Message Service | View ] To View Them." & vbCrLf
                                If U1.ListItems.Item(jo).SubItems(2) = "1" Then
                                    Call sendall("JOIN|¿|" & nick(Index))
                                    Call sendall("MSG|¿|SERVER >> User " & nick(Index) & " is set as a Moderator." & vbCrLf)
                                Else
                                    Call sendall("JOIN|¿|" & nick(Index))
                                End If
                                RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & split1(1) & " Joined Room..."
                                Exit Sub
                        End If
                    End If
                Next
                Winsock1(Index).SendData "INVALID|¿|669" & vbCrLf
                
                'CHAT MSG SENT
            Case "MSG"
            'Muzzle Check
            Dim opers As String
            
            If Left(UCase(split1(1)), 4) = "/OPS" Then
            opers = "Current Online Ops : "
                For jo = 1 To U2.ListItems.Count
                    If U2.ListItems.Item(jo).SubItems(1) = "1" Then
                    opers = opers & U2.ListItems.Item(jo) & " - "
                    End If
                Next
                Winsock1(Index).SendData "SMSG|¿|" & opers
                Exit Sub
            End If
                If Left(UCase(split1(1)), 4) = "/OPP" Then
                opers = "Current Registered Ops : "
                For jo = 1 To U1.ListItems.Count
                    If U1.ListItems.Item(jo).SubItems(2) = "1" Then
                    opers = opers & U1.ListItems.Item(jo) & " - "
                    End If
                Next
                Winsock1(Index).SendData "SMSG|¿|" & opers
                Exit Sub
            End If
            On Error Resume Next
                For jo = 1 To M1.ListItems.Count
                    If nick(Index) = M1.ListItems.Item(jo) Then
                        Winsock1(Index).SendData "SMSG|¿| You Are Muzzled For " & M1.ListItems.Item(jo).SubItems(1) & " More Seconds..."
                        Exit Sub
                    End If
                Next

            'Chr() Removal
            split1(1) = Replace(split1(1), "                                   ", " ")
            For re = 0 To 20
                DoEvents
                split1(1) = Replace(split1(1), Chr(re), "")
nexter76:
            Next
            For re = 150 To 255
                DoEvents
                split1(1) = Replace(split1(1), Chr(re), "")
            Next
            If Len(split1(1)) > 201 Then
            On Error Resume Next
            Winsock1(Index).SendData "SMSG|¿| To Long Over 200 chr's"
            GoTo Youknow:
            End If
            If Len(nick(Index)) < 1 Then
            On Error Resume Next
            Winsock1(Index).Close
            GoTo Youknow:
            End If
        
            On Error Resume Next
                If nick(Index) = "" Then
                On Error Resume Next
                    Winsock1(Index).Close
                    Call Winsock1_Close(Index)
                End If
                            '''Swear Filter
            Dim word As Variant
            For re = 0 To List8.ListCount
                If InStr(LCase(split1(1)), LCase(List8.List(re))) Then
                    word = Split(split1(1), " ")
                    split1(1) = ""
                    For jo = 0 To UBound(word)
                        If LCase(word(jo)) = LCase(List8.List(re)) Then
                            word(jo) = Replace(LCase(word(jo)), LCase(List8.List(re)), "****")
                        End If
                        split1(1) = split1(1) & word(jo) & " "
                    Next
                End If
            Next
                If UCase(Left(split1(1), 5)) = "!SEEN" Then
                    Dim seen As Variant
                    On Error Resume Next
                    seen = Split(split1(1), " ")
Dim reop As String
                            For re = 1 To U1.ListItems.Count
                            On Error Resume Next
                            If LCase(seen(1)) = LCase(U1.ListItems.Item(re)) Then
                                reop = U1.ListItems.Item(re).SubItems(5) & " At " & U1.ListItems.Item(re).SubItems(4) & " US Eastern Time."
                            End If
                            Next
                    For re = 1 To U2.ListItems.Count
                        If LCase(seen(1)) = LCase(U2.ListItems.Item(re)) Then
                            Call sendall("MSG|¿|SERVER >> User " & seen(1) & " is Currently online. And Logged in " & reop)
                            Exit Sub
                        End If
                    Next
                    For re = 1 To U1.ListItems.Count
                        If LCase(seen(1)) = LCase(U1.ListItems.Item(re)) Then
                            Call sendall("MSG|¿|SERVER >> User " & seen(1) & " Was Last Seen " & U1.ListItems.Item(re).SubItems(5) & " At " & U1.ListItems.Item(re).SubItems(4) & " US Eastern Time.")
                            Exit Sub
                        End If
                    Next
                Call sendall("MSG|¿|SERVER >> User " & seen(1) & " Is Not a Registered User.")
                Exit Sub
                End If
                Call sendall("MSG|¿|<" & nick(Index) & "> " & split1(1))
                On Error Resume Next
                RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & "<" & nick(Index) & "> " & split1(1)
'FLOODCATCH
                    If LCase(nick(Index)) = "vbbot" Or LCase(nick(Index)) = "Timothy" Then
                        List1.Clear
                        List2.Clear
                        GoTo kooloo:
                    End If
                    If List1.List(0) = "<" & nick(Index) & "> " & split1(1) Then
                        List1.AddItem "<" & nick(Index) & "> " & split1(1)
                        If List1.ListCount > 2 Then
                            List3.AddItem Winsock1(Index).RemoteHostIP
                                                                MS1.ListItems.Add , , nick(Index)
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(1) = "SERVERÞ"
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(2) = Date & " : " & Time
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(3) = "YOU WHERE KICKED : REPEATING!"
                            jh = tally(Winsock1(Index).RemoteHostIP)
                            Call Winsock1_Close(Index)
                            Winsock1(Index).Close
                            List1.Clear
                            List2.Clear
                            GoTo Youknow:
                        End If
                        Else
                        List1.Clear
                        List1.AddItem "<" & nick(Index) & "> " & split1(1)
                        End If
                        
                        If List2.List(0) = Winsock1(Index).RemoteHostIP Then
                        List2.AddItem Winsock1(Index).RemoteHostIP
                        If List2.ListCount > 5 Then
                            List3.AddItem Winsock1(Index).RemoteHostIP
                                                                                        MS1.ListItems.Add , , nick(Index)
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(1) = "SERVERÞ"
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(2) = Date & " : " & Time
                                    MS1.ListItems.Item(MS1.ListItems.Count).SubItems(3) = "YOU WHERE KICKED : FLOOD DETECTED!"
                            jh = tally(Winsock1(Index).RemoteHostIP)
                            Call Winsock1_Close(Index)
                            Winsock1(Index).Close
                            List2.Clear
                            List1.Clear
                            GoTo Youknow:
                        End If
                        Else
                        List2.Clear
                        List2.AddItem Winsock1(Index).RemoteHostIP
                        End If
kooloo:
'Op Commands
            Case "OPS"
            Dim jol As Integer
                For jol = 1 To U2.ListItems.Count
                    If nick(Index) = U2.ListItems.Item(jol) Then
                        If U2.ListItems.Item(jol).SubItems(1) = "1" Then
                            GoTo ISOP
                            
                        End If
                    End If
                Next
                If Winsock1(Index).State = 7 Then
                    Winsock1(Index).SendData "SMSG|¿|You Are Not Recognized as an Op."
                End If
                Exit Sub
ISOP:
        If split1(1) = "CPB" Then
            List5.Clear
            Call sendall("MSG|¿|SERVER >> Permanent Bans Cleared")
        Exit Sub
        End If
        If split1(1) = "CRB" Then
            List4.Clear
            Call sendall("MSG|¿|SERVER >> 60 Min Bans Cleared")
        Exit Sub
        End If
        If split1(1) = "CSB" Then
            List7.Clear
            Call sendall("MSG|¿|SERVER >> Subnet Bans Cleared")
        Exit Sub
        End If
    
                If split1(1) = "CSB" Then
            List7.Clear
            Call sendall("MSG|¿|SERVER >> Subnet Bans Cleared")
        Exit Sub
        End If
        
        
        
        
        If split1(1) = "GBL" Then
        On Error Resume Next
                 Dim banners As String
                 banners = "60 Min Bans : "
                For jo = 0 To List4.ListCount
                     banners = banners & List4.List(jo) & " - "
                 Next
                Winsock1(Index).SendData "SMSG|¿|" & banners
                Exit Sub

        End If
        
        
        If split1(1) = "GPBL" Then
         On Error Resume Next
                 Dim banner As String
                 banner = "Perament Bans List : "
                For jo = 0 To List5.ListCount
                     banner = banner & List5.List(jo) & " - "
                 Next
                Winsock1(Index).SendData "SMSG|¿|" & banner
                Exit Sub

        End If
        If split1(1) = "GSBL" Then
         On Error Resume Next
                 Dim banne As String
                 banne = "SubNet Ban List : "
                For jo = 0 To List7.ListCount
                     banne = banne & List7.List(jo) & " - "
                 Next
                Winsock1(Index).SendData "SMSG|¿|" & banne
                Exit Sub

        End If
        If UBound(split1) < 2 Then Exit Sub
        On Error Resume Next
                Dim IU2 As Integer
                For jol = 1 To U2.ListItems.Count
                    If split1(2) = U2.ListItems.Item(jol) Then
                    IU2 = jol
                    GoTo GOTINDEX
                    End If
                Next
                If Winsock1(Index).State = 7 Then
                    Winsock1(Index).SendData "SMSG|¿|Error Completeing Op Command"
                End If
                Exit Sub
GOTINDEX:
                Select Case split1(1)
                    Case "KICK"
                        On Error Resume Next
                        If U2.ListItems.Item(IU2).SubItems(1) > 0 Then
                            Winsock1(Index).SendData "SMSG|¿|Cant Kill Other Ops."
                        Exit Sub
                        End If
                            Winsock1_Close U2.ListItems.Item(IU2).SubItems(2)
                            Exit Sub
                    Case "BAN"
                        On Error Resume Next
                        If U2.ListItems.Item(IU2).SubItems(1) > 0 Then
                            Winsock1(Index).SendData "SMSG|¿|Cant Ban Other Ops."
                        Exit Sub
                        End If
                            List4.AddItem Winsock1(U2.ListItems.Item(IU2).SubItems(2)).RemoteHostIP
                        Call sendall("MSG|¿|SERVER >> " & U2.ListItems.Item(IU2) & " Has Been Banned.")
                    Case "VIEW"
                         On Error Resume Next
                         If Winsock1(Index).State = 7 Then
                            Winsock1(Index).SendData "SMSG|¿|" & U2.ListItems.Item(IU2) & "'s Ip Is " & U2.ListItems.Item(IU2).SubItems(3)
                         End If
                    Case "NAME"
                        On Error Resume Next
                        If U2.ListItems.Item(IU2).SubItems(1) > 0 Then
                            Winsock1(Index).SendData "SMSG|¿|Cant Ban Other Ops."
                        Exit Sub
                        End If
                            List4.AddItem nick(U2.ListItems.Item(IU2).SubItems(2))
                        Call sendall("MSG|¿|SERVER >> " & U2.ListItems.Item(IU2) & "'s Nick Has Been Banned.")
                        
                    Case "SUBNET"
                        On Error Resume Next
                        If U2.ListItems.Item(IU2).SubItems(1) > 0 Then
                            Winsock1(Index).SendData "SMSG|¿|Cant Ban Other Ops."
                        Exit Sub
                        End If
                            Dim subn As Variant
                            Dim subn1 As String
                            subn = Split(U2.ListItems.Item(IU2).SubItems(3), ".")
                            subn1 = subn(0) & "." & subn(1)
                            List7.AddItem subn1
                        Call sendall("MSG|¿|SERVER >> " & U2.ListItems.Item(IU2) & "'s Subnet Has Been Banned (65025 ip's)")
                    Case "MUZZLE"
                    
                        If U2.ListItems.Item(IU2).SubItems(1) > 0 Then
                            Winsock1(Index).SendData "SMSG|¿|Cant Muzzle Other Ops."
                        Exit Sub
                        End If
                        On Error Resume Next
                        M1.ListItems.Add , , U2.ListItems.Item(IU2)
                        If UBound(split1) > 2 Then
                            If IsNumeric(split1(3)) = True Then
                            If split1(3) > (20 * 60) Then GoTo morethan20
                                M1.ListItems.Item(M1.ListItems.Count).SubItems(1) = split1(3)
                                Call sendall("MSG|¿|SERVER >> " & U2.ListItems.Item(IU2) & " Has Been Muzzled For " & (split1(3) / 60) & " Min(" & split1(3) & " Seconds).")
                                Exit Sub
                            End If
                        End If
morethan20:
                        M1.ListItems.Item(M1.ListItems.Count).SubItems(1) = "300"
                        Call sendall("MSG|¿|SERVER >> " & U2.ListItems.Item(IU2) & " Has Been Muzzled For 5 Min(300 Seconds).")
                End Select
            Case "PMSG"
                On Error Resume Next
                For jo = 1 To M1.ListItems.Count
                    If nick(Index) = M1.ListItems.Item(jo) Then
                        Winsock1(Index).SendData "SMSG|¿| You Are Muzzled For " & M1.ListItems.Item(jo).SubItems(1) & " More Seconds..."
                        Exit Sub
                    End If
                Next
            If UBound(split1) < 2 Then GoTo Youknow:
                If Len(nick(Index)) < 1 Then
                Winsock1(Index).Close
                End If
                For i = 1 To Winsock1.UBound
                    If nick(i) = split1(1) Then
                        If InStr(split1(2), "/SENDFILE") Then
                        On Error Resume Next
                            Winsock1(i).SendData "PMSG|¿|<" & nick(Index) & "> " & split1(2) & "|@|" & Winsock1(Index).RemoteHostIP & vbCrLf
                            GoTo Youknow:
                        Else
                            On Error Resume Next
                            Winsock1(i).SendData "PMSG|¿|<" & nick(Index) & "> " & split1(2) & vbCrLf
                            GoTo Youknow:
                        End If
                    End If
                Next
                Winsock1(Index).SendData "SMSG|¿|" & split1(1) & " : Is Not Currently Online."
Youknow:
        End Select
    Next
End Sub

Sub sendall(shimi As String)
Dim wt As Integer
    If wt = 1 Then
        Do Until wt = 0
            DoEvents
        Loop
    End If
    wt = 1
    On Error Resume Next
    For i = 1 To U2.ListItems.Count
        On Error Resume Next
        If Winsock1(U2.ListItems.Item(i).SubItems(2)).State = 7 Then
            On Error Resume Next
            Winsock1(U2.ListItems.Item(i).SubItems(2)).SendData shimi & vbCrLf
        End If
    Next
    wt = 0
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
For i = 1 To P1.ListItems.Count
On Error Resume Next
    If P1.ListItems.Item(i) = Index Then
        P1.ListItems.Remove i
    End If
Next
For i = 1 To LV1.ListItems.Count
On Error Resume Next
    If LV1.ListItems.Item(i) = nick(Index) Then
    LV1.ListItems.Remove i
    End If
Next
For i = 1 To U2.ListItems.Count
On Error Resume Next
    If U2.ListItems.Item(i) = nick(Index) Then
        Dim vg
        For vg = 1 To U1.ListItems.Count
            If U1.ListItems.Item(vg) = nick(Index) Then
                U1.ListItems.Item(vg).SubItems(4) = Time
                U1.ListItems.Item(vg).SubItems(5) = Date
            End If
        Next
    On Error Resume Next
    U2.ListItems.Remove i
    End If
Next
done:
RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & Winsock1(Index).RemoteHostIP & " Disconnected...(" & Index & ")" & vbCrLf
RTB.Text = RTB.Text & vbCrLf & "[" & Time & "] " & nick(Index) & " Left...(" & Index & ")" & vbCrLf
On Error Resume Next
If nick(Index) = "" Then
Winsock1(Index).Close
bog(Index) = 0
Exit Sub
End If
On Error Resume Next
Call sendall("PART|¿|" & nick(Index))
nick(Index) = ""
bog(Index) = 0
Winsock1(Index).Close
End Sub

Public Function tally(ip As String)
    Dim foo As Integer
    foo = 0
    For i = 0 To List3.ListCount - 1
        If List3.List(i) = ip Then
            foo = foo + 1
        End If
    Next
    tally = foo
    If foo > Text7.Text - 1 Then
        List4.AddItem ip
        For i = 0 To List3.ListCount - 1
            If List3.List(i) = ip Then
                List3.RemoveItem i
            End If
        Next
    End If
End Function

