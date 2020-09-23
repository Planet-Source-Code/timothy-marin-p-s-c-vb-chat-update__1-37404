Attribute VB_Name = "Module1"
Option Explicit
Public pm(50) As New Form3
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long


 
Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400
Public Const WM_NOTIFY = &H4E
Public Const WM_LBUTTONDOWN = &H201
Public Const EM_GETEVENTMASK = WM_USER + 59
Public Const EM_GETTEXTRANGE = WM_USER + 75

Public Const EM_SETEVENTMASK = WM_USER + 69
Public Const EM_AUTOURLDETECT = WM_USER + 91
Public Const EN_LINK = &H70B

Public Const ENM_LINK = &H4000000
Public Const SW_SHOWNORMAL = 1

Type tagNMHDR
    hwndFrom As Long
    idFrom   As Long
    code     As Long
End Type

Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Type ENLINK
    nmhdr  As tagNMHDR
    msg    As Long
    wParam As Long
    lParam As Long
    chrg   As CHARRANGE
End Type

Type TEXTRANGE
    chrg      As CHARRANGE
    lpstrText As Long
End Type

Public glnglpOriginalWndProc As Long
Public glngOriginalhWnd As Long

Function RichTextBoxSubProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim udtNMHDR               As tagNMHDR
Dim udtENLINK              As ENLINK
Dim udtTEXTRANGE           As TEXTRANGE
Dim strBuffer              As String * 128
Dim strOperation           As String
Dim strFileName            As String
Dim strDefaultDirectory    As String
Dim lngHInstanceExecutable As Long
Dim lngWin32apiResultCode  As Long
 
 
 If uMsg = WM_NOTIFY Then
    RtlMoveMemory udtNMHDR, ByVal lParam, Len(udtNMHDR)
    If udtNMHDR.hwndFrom = Form1.rtb.hWnd And udtNMHDR.code = EN_LINK Then
        RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
        If udtENLINK.msg = WM_LBUTTONDOWN Then
            strBuffer = ""
            
            With udtTEXTRANGE
                .chrg.cpMin = udtENLINK.chrg.cpMin
                .chrg.cpMax = udtENLINK.chrg.cpMax
                .lpstrText = StrPtr(strBuffer)
            End With
 
            With Form1.rtb
                lngWin32apiResultCode = SendMessage(.hWnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
            End With

            RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
            strOperation = "open"
            strFileName = strBuffer
            lngHInstanceExecutable = ShellExecute(Form1.hWnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)

        End If
    End If
End If
  
RichTextBoxSubProc = CallWindowProc(glnglpOriginalWndProc, hWnd, uMsg, wParam, lParam)

End Function
Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
' Example: Call FormOnTop(me.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Dim Placement As Long
    'wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, SWP_NOSIZE
End Sub
