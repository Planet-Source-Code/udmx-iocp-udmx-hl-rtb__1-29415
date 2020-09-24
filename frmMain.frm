VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00B07315&
   Caption         =   "UDMX_HL_RTB_EDITOR"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   635
      ButtonWidth     =   5927
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select Text Color < SelColor >"
            ImageKey        =   "B_AO"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select Text Highlight Color < SelHColor >"
            ImageKey        =   "R_AO"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select Font Name < SelFontName >"
            ImageKey        =   "Font"
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4545
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18018
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/5/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "9:58 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2160
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   "File"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101C
            Key             =   "Bullet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":136E
            Key             =   "R_AO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16C0
            Key             =   "B_AO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A12
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D64
            Key             =   "Open1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20B6
            Key             =   "Save1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2408
            Key             =   "Undo1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":275A
            Key             =   "Redo1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AAC
            Key             =   "Cut1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DFE
            Key             =   "Copy1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3150
            Key             =   "Paste1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34A2
            Key             =   "Delete1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37F4
            Key             =   "Italic1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B46
            Key             =   "Bold1"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E98
            Key             =   "Underline1"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41EA
            Key             =   "StrikeThough"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "File"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo1"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut1"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy1"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste1"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete1"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic1"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold1"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline1"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strike Through"
            Object.ToolTipText     =   "Strike Through"
            ImageKey        =   "StrikeThough"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullet"
            ImageKey        =   "Bullet"
         EndProperty
      EndProperty
      Begin VB.ComboBox ComboFontSize 
         Height          =   315
         ItemData        =   "frmMain.frx":453C
         Left            =   6600
         List            =   "frmMain.frx":455B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Text            =   "Font Size:"
         Top             =   50
         Width           =   975
      End
      Begin VB.CommandButton cmd_ChangePicture 
         Caption         =   "Change Current Form's Picture"
         Height          =   350
         Left            =   8400
         TabIndex        =   3
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a small Editor I made to test out the OCX
'Please vote for this if you like it..
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7

Private Sub cmd_ChangePicture_Click()
On Error Resume Next
With dlgCommonDialog
    .ShowOpen
    .Filter = "Pictures(*.bmp;*.ico;*.jpg;*.jpeg)|*.bmp;*.ico;*.jpg;*.jpeg|All Files (*.*)|*.*"
    ActiveForm.RTFTEXT.Picture = LoadPicture(.FileName)
End With
End Sub

Private Sub ComboFontSize_Click()
    ActiveForm.RTFTEXT.SelFontSize = ComboFontSize.Text

End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000) 'load from Setting Registry
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc 'A function that makes a new Untitled Document
End Sub


Private Sub LoadNewDoc() 'A function that makes a new Untitled Document
Dim answer As String

    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
    answer = MsgBox("Do you want to Set a new background picture for this new file?", vbYesNoCancel, "Set Picture")
    If answer = vbCancel Then Unload frmD
    If answer = vbYes Then
        On Error Resume Next
        With dlgCommonDialog
            .ShowOpen
            .Filter = "Pictures(*.bmp;*.ico;*.jpg;*.jpeg)|*.bmp;*.ico;*.jpg;*.jpeg|All Files (*.*)|*.*"
            ActiveForm.RTFTEXT.Picture = LoadPicture(.FileName)
        End With
    End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then 'save to Registry
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEditRedo_Click()
    If ActiveForm.Redo = False Then
        tbToolBar.Buttons(6).Enabled = False
        mnuEditRedo.Enabled = False
    End If
    mnuEditUndo.Enabled = True
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            Dim newdocument As New frmDocument
            newdocument.Show
            dlgCommonDialog.Filter = ""
            dlgCommonDialog.ShowOpen
            ActiveForm.RTFTEXT.FileName = dlgCommonDialog.FileName
        Case "Undo"
            If ActiveForm.Undo = False Then
                Button.Enabled = False
                mnuEditUndo.Enabled = False
            End If
        Case "Redo"
            If ActiveForm.Redo = False Then
                Button.Enabled = False
                mnuEditRedo.Enabled = False
            End If
        Case "Delete"
            ActiveForm.RTFTEXT.SelText = ""
        Case "New"
            ActiveForm.RTFTEXT.TextRTF = ""
        Case "Strike Through"
            ActiveForm.RTFTEXT.SelStrikeThru = Not ActiveForm.RTFTEXT.SelStrikeThru
            Button.Value = IIf(ActiveForm.RTFTEXT.SelStrikeThru, tbrPressed, tbrUnpressed)
        Case "Save"
            dlgCommonDialog.Filter = ""
            dlgCommonDialog.FileName = ActiveForm.Caption
            dlgCommonDialog.ShowSave
            On Error Resume Next
            ActiveForm.RTFTEXT.SaveFile dlgCommonDialog.FileName
        Case "Bullet"
            ActiveForm.RTFTEXT.SelBullet = Not ActiveForm.RTFTEXT.SelBullet
            Button.Value = IIf(ActiveForm.RTFTEXT.SelBullet, tbrPressed, tbrUnpressed)
        Case "Cut"
            Clipboard.SetText ActiveForm.RTFTEXT.SelRTF, vbCFRTF
            ActiveForm.RTFTEXT.SelText = ""
        Case "Copy"
            Clipboard.SetText ActiveForm.RTFTEXT.SelRTF, vbCFRTF
        Case "Paste"
            ActiveForm.RTFTEXT.SelRTF = Clipboard.GetText(vbCFRTF)
        Case "Bold"
            ActiveForm.RTFTEXT.SelBold = Not ActiveForm.RTFTEXT.SelBold
            Button.Value = IIf(ActiveForm.RTFTEXT.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.RTFTEXT.SelItalic = Not ActiveForm.RTFTEXT.SelItalic
            Button.Value = IIf(ActiveForm.RTFTEXT.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.RTFTEXT.SelUnderline = Not ActiveForm.RTFTEXT.SelUnderline
            Button.Value = IIf(ActiveForm.RTFTEXT.SelUnderline, tbrPressed, tbrUnpressed)
    End Select

End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.RTFTEXT.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.RTFTEXT.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.RTFTEXT.SelRTF
    ActiveForm.RTFTEXT.SelText = vbNullString

End Sub


Private Sub mnuEditUndo_Click()
    If ActiveForm.Undo = False Then
        tbToolBar.Buttons(5).Enabled = False
        mnuEditUndo.Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case "1"
            On Error Resume Next
            Me.dlgCommonDialog.Color = ActiveForm.RTFTEXT.SelColor
            Me.dlgCommonDialog.ShowColor
            ActiveForm.RTFTEXT.SelColor = Me.dlgCommonDialog.Color
        Case "2"
            On Error Resume Next
            Me.dlgCommonDialog.Color = Me.ActiveForm.RTFTEXT.SELHCOLOR
            Me.dlgCommonDialog.ShowColor
            ActiveForm.RTFTEXT.SELHCOLOR = Me.dlgCommonDialog.Color

    End Select
End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If Button.Index = 3 Then
        Dim IX As Integer
        For IX% = 0 To Screen.FontCount - 1
            Button.ButtonMenus.Add , , Screen.Fonts(IX%)
        Next IX%
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    ActiveForm.RTFTEXT.SelFontName = ButtonMenu.Text
End Sub

