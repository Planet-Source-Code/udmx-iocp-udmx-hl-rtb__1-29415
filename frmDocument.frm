VERSION 5.00
Object = "*\A..\..\..\..\..\Program Files\Microsoft Visual Studio\VB98\Secret OCX\URTB\UDMX HL RTB.VBP"
Begin VB.Form frmDocument 
   Caption         =   "Untitled"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   6735
   Begin UDMX_HL_RTB.UDRichTextBox RTFtext 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDocument.frx":030A
      Text            =   ""
      MouseIcon       =   "frmDocument.frx":10615
      ScrollBars      =   2
      Transparent_RTF =   -1  'True
      SelRTF          =   $"frmDocument.frx":10631
      SelFontSize     =   8.25
      SelFontName     =   "MS Sans Serif"
      TextRTF         =   $"frmDocument.frx":10666
      SelHColor       =   0
      SelFontName     =   ""
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'You can use this code <RTB UNDO REDO> 100% free < No need to Put my name on your Program >
Dim MakeChange As Boolean
Dim CHistoryNumber As Integer
Dim History_Box(5000) As String
Dim MaxHistoryNumber As Integer
Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    RTFTEXT.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - 100
    RTFTEXT.RightMargin = RTFTEXT.Width - 400
End Sub
Public Function Undo() As Boolean
    MakeChange = True
    If CHistoryNumber <= 1 Then
        Undo = False
    Else
        Undo = True
    End If
    CHistoryNumber = CHistoryNumber - 1
    RTFTEXT.TextRTF = History_Box(CHistoryNumber)
    fMainForm.tbToolBar.Buttons("Redo").Enabled = True
    fMainForm.tbToolBar.Buttons("Undo").Enabled = True
    fMainForm.mnuEditRedo.Enabled = True
    MakeChange = False
End Function
Public Function Redo() As Boolean
    MakeChange = True
    If CHistoryNumber >= MaxHistoryNumber Then
        Redo = False
    Else
        Redo = True
    End If
    CHistoryNumber = CHistoryNumber + 1
    If CHistoryNumber = MaxHistoryNumber Then Redo = False
    RTFTEXT.TextRTF = History_Box(CHistoryNumber)
    fMainForm.tbToolBar.Buttons("Undo").Enabled = True
    fMainForm.mnuEditRedo.Enabled = True
    MakeChange = False
End Function

Private Sub RTFTEXT_Change()
If Not MakeChange Then
    fMainForm.tbToolBar.Buttons("Undo").Enabled = True
    fMainForm.mnuEditUndo.Enabled = True
    MaxHistoryNumber = CHistoryNumber + 1
    CHistoryNumber = CHistoryNumber + 1
    History_Box(CHistoryNumber) = RTFTEXT.TextRTF
End If
End Sub

