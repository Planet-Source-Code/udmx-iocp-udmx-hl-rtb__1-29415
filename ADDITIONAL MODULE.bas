Attribute VB_Name = "TRANSPARENT_HIGHLIGHT_MODULES"
Option Explicit
Public Const WM_USER = &H400
Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000

Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Const SW_SHOWNORMAL = 1
Dim htxt As String
Dim Word As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim Mtoolbar_GF As Boolean
Dim outpf As Boolean
Dim mesgf As Boolean
Dim g_shifted As Boolean
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(5000) As String
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Public Function MakeRTBTransparent(RTBCtl As Object) As Boolean

On Error Resume Next
RTBCtl.BackColor = RTBCtl.Parent.BackColor
SetWindowLong RTBCtl.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
MakeRTBTransparent = Err.LastDllError = 0

End Function


