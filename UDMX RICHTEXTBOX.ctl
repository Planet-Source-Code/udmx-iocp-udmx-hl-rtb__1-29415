VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl UDRichTextBox 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ScaleHeight     =   3975
   ScaleWidth      =   7365
   ToolboxBitmap   =   "UDMX RICHTEXTBOX.ctx":0000
   Begin RichTextLib.RichTextBox URTB 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"UDMX RICHTEXTBOX.ctx":0312
   End
End
Attribute VB_Name = "UDRichTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This OCX + the Source code is in fact "Copyrighted" however with an exception
' You may use this ocx + source code in your program as long as you mention my nick name, UDMX IOCP, in your program.
' In addition to this I will create other active-x and other misc code if you vote 4 me. :) hope to win....
'UDMX IOCP : Bringing you the source"
Const m_def_SelFontName = "0"
Const m_def_SelHColor = &HFFFFFF
Dim udtCharFormat As CHARFORMAT2
Const m_def_SelUnderline = 0
Const m_def_SelText = ""
Const m_def_SelTabCount = 0
Const m_def_SelStrikeThru = 0
Const m_def_SelStart = 0
Const m_def_SelRTF = "0"
Const m_def_SelRightIndent = 0
Const m_def_SelProtected = 0
Const m_def_SelLength = 0
Const m_def_SelItalic = 0
Const m_def_SelIndent = 0
Const m_def_SelHangingIndent = 0
Const m_def_SelFontSize = 0
Const m_def_TextRTF = 0
Const m_def_SelAlignment = 0
Const m_def_SelBold = 0
Const m_def_SelBullet = 0
Const m_def_SelCharOffset = 0
Const m_def_SelColor = 0
Const m_def_Transparent_RTF = 0
Dim m_SelFontName As String
Dim m_SelHColor As OLE_COLOR
Dim m_SelUnderline As Boolean
Dim m_SelText As String
Dim m_SelTabCount As Integer
Dim m_SelStrikeThru As Boolean
Dim m_SelStart As Integer
Dim m_SelRTF As String
Dim m_SelRightIndent As Integer
Dim m_SelProtected As Boolean
Dim m_SelLength As Integer
Dim m_SelItalic As Boolean
Dim m_SelIndent As Integer
Dim m_SelHangingIndent As Integer
Dim m_SelFontSize As Integer

Dim m_TextRTF As String
Dim m_SelAlignment As Variant
Dim m_SelBold As Boolean
Dim m_SelBullet As Boolean
Dim m_SelCharOffset As Boolean
Dim m_SelColor As OLE_COLOR
Public Enum gga
    rtfNoBorder = 0
    rtfFixedSingle = 1
End Enum
Public Enum ggb
    rtfThreeD = 1
    Flat = 0
End Enum
Dim m_Transparent_RTF As Boolean

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event SelChange()
Event Validate(Cancel As Boolean)
Const SW_SHOWNORMAL = 1

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
On Error Resume Next
    BackColor = URTB.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    URTB.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
On Error Resume Next
    Enabled = URTB.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    URTB.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = URTB.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set URTB.Font = New_Font
    PropertyChanged "Font"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
    URTB.Refresh
End Sub

Private Sub URTB_Click()
    RaiseEvent Click

End Sub


Private Sub URTB_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub URTB_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub URTB_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub URTB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub URTB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub URTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Public Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
On Error Resume Next
    AutoVerbMenu = URTB.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
    URTB.AutoVerbMenu() = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

Public Property Get BulletIndent() As Single
Attribute BulletIndent.VB_Description = "Returns or sets the amount of indent used in a RichTextBox control when SelBullet is set to True."
On Error Resume Next
    BulletIndent = URTB.BulletIndent
End Property

Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
    URTB.BulletIndent() = New_BulletIndent
    PropertyChanged "BulletIndent"
End Property

Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
On Error Resume Next
    CausesValidation = URTB.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    URTB.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property

Private Sub URTB_Change()
    RaiseEvent Change
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
On Error Resume Next
    hWnd = UserControl.hWnd
End Property

Public Sub LoadFile(ByVal bstrFilename As String, Optional ByVal vFileType As Variant)
Attribute LoadFile.VB_Description = "Loads an .RTF file or text file into a RichTextBox control."
    URTB.LoadFile bstrFilename, vFileType
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
On Error Resume Next
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub URTB_SelChange()
    RaiseEvent SelChange
End Sub

Public Sub SelPrint(ByVal lHDC As Long, Optional ByVal vStartDoc As Variant)
Attribute SelPrint.VB_Description = "Sends formatted text in a RichTextBox control to a device for printing."
    URTB.SelPrint lHDC, vStartDoc
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
On Error Resume Next
    Text = URTB.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    URTB.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the filename of the file loaded into the RichTextBox control at design time."
On Error Resume Next
    FileName = URTB.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    URTB.FileName() = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that specifies if the selected item remains highlighted when a control loses focus."
On Error Resume Next
    HideSelection = URTB.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    URTB.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets a value indicating whether there is a maximum number of characters a RichTextBox control can hold and, if so, specifies the maximum number of characters."
On Error Resume Next
    MaxLength = URTB.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    URTB.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
On Error Resume Next
    Set MouseIcon = URTB.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set URTB.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets a value indicating the type of mouse pointer displayed when the mouse is over the control at run time."
On Error Resume Next
    MousePointer = URTB.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    URTB.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
On Error Resume Next
    RightMargin = URTB.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
    URTB.RightMargin() = New_RightMargin
    PropertyChanged "RightMargin"
End Property

Public Property Get ScrollBars() As ScrollBarsConstants
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether a RichTextBox control has horizontal or vertical scroll bars."
On Error Resume Next
    ScrollBars = URTB.ScrollBars
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
On Error Resume Next
    WhatsThisHelpID = URTB.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    URTB.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
On Error Resume Next
    ToolTipText = URTB.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    URTB.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub URTB_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Integer = 0, Optional ByVal vOptions As Variant) As Long
Attribute Find.VB_Description = "Searches the text in a RichTextBox control for a given string."
    Dim ax, am, aad, ddm
    ax = Len(URTB.Text)
    If vEnd = 0 Then vEnd = ax
    If URTB.SelStart > -1 Then vStart = URTB.SelStart + URTB.SelLength
    If URTB.SelStart = ax Then vStart = 0
    For am = vStart To vEnd
        aad = Len(bstrString)

        URTB.SelStart = am
        URTB.SelLength = aad
        Select Case vOptions
            Case 4
                If URTB.SelText = bstrString Then
                    Find = 1
                    Exit Function
                End If
            Case 0
                If LCase(URTB.SelText) = LCase(bstrString) Then
                    Find = 1
                    Exit Function
                End If
            Case 2
                If am <= 0 Then
                    Dim AMA
                    
                    AMA = Left$(URTB.Text, aad + 1)
                    If LCase(AMA) = LCase(bstrString & " ") Then
                        Find = 1
                        Exit Function
                    End If
                ElseIf am + aad >= ax Then
                    Dim AXA
                    AXA = Right$(URTB.Text, aad + 1)
                    If LCase(AXA) = " " & LCase(bstrString) Then
                        Find = 1
                        Exit Function
                    End If
                Else
                    If Mid(URTB.Text, am, aad + 2) = " " & bstrString & " " Then
                        Find = 1
                        Exit Function
                    End If
                End If
            Case 6
                If am <= 0 Then
                    Dim AMAb
                    
                    AMAb = Left$(URTB.Text, aad + 1)
                    If AMAb = bstrString & " " Then
                        Find = 1
                        Exit Function
                    End If
                ElseIf am + aad >= ax Then
                    Dim AXAb
                    AXAb = Right$(URTB.Text, aad + 1)
                    If AXAb = " " & bstrString Then
                        Find = 1
                        Exit Function
                    End If
                Else
                    If Mid(URTB.Text, am, aad + 2) = " " & bstrString & " " Then
                        Find = 1
                        Exit Function
                    End If
                End If

        End Select
    Next am
    Find = -1
End Function

Public Function GetLineFromChar(ByVal lChar As Long) As Long
Attribute GetLineFromChar.VB_Description = "Returns the number of the line containing a specified character position in a RichTextBox control."
    GetLineFromChar = URTB.GetLineFromChar(lChar)
End Function

Public Sub SaveFile(ByVal bstrFilename As String, Optional ByVal vFlags As Variant = "")
Attribute SaveFile.VB_Description = "Saves the contents of a RichTextBox control to a file."

    If vFlags <> "" Then
        URTB.SaveFile bstrFilename, vFlags
    Else
        URTB.SaveFile bstrFilename
    End If
End Sub

Public Sub Span(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute Span.VB_Description = "Selects text in a RichTextBox control based on a set of specified characters."
    URTB.Span bstrCharacterSet, vForward, vNegate
End Sub

Public Sub UpTo(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute UpTo.VB_Description = "Moves the insertion point up to, but not including, the first character that is a member of the specified character set in a RichTextBox control."
    URTB.UpTo bstrCharacterSet, vForward, vNegate
End Sub

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents in a RichTextBox control can be edited."
On Error Resume Next
    Locked = URTB.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    URTB.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get Transparent_RTF() As Boolean
Attribute Transparent_RTF.VB_Description = "Return or set transparency of the currently selected text in a RichTextBox control. Not available at design time."
On Error Resume Next
    Transparent_RTF = m_Transparent_RTF
    UserControl.BackStyle = 0
End Property

Public Property Let Transparent_RTF(ByVal New_Transparent_RTF As Boolean)
    m_Transparent_RTF = New_Transparent_RTF
    If m_Transparent_RTF = True Then
        MakeRTBTransparent URTB
        UserControl.BackStyle = 0
    End If
    PropertyChanged "Transparent_RTF"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    URTB.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    URTB.Enabled = PropBag.ReadProperty("Enabled", True)
    Set URTB.Font = PropBag.ReadProperty("Font", Ambient.Font)
    URTB.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    URTB.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    URTB.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    URTB.Text = PropBag.ReadProperty("Text", "RichTextBox1")
    URTB.FileName = PropBag.ReadProperty("FileName", "")
    URTB.HideSelection = PropBag.ReadProperty("HideSelection", True)
    URTB.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    URTB.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    URTB.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    URTB.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    URTB.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    URTB.Locked = PropBag.ReadProperty("Locked", False)
    m_Transparent_RTF = PropBag.ReadProperty("Transparent_RTF", False)
    If m_Transparent_RTF = True Then MakeRTBTransparent URTB
    UserControl.BackStyle = 0
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    URTB.SelUnderline = PropBag.ReadProperty("SelUnderline", m_def_SelUnderline)
    URTB.SelText = PropBag.ReadProperty("SelText", m_def_SelText)
    URTB.SelTabCount = PropBag.ReadProperty("SelTabCount", m_def_SelTabCount)
    URTB.SelStrikeThru = PropBag.ReadProperty("SelStrikeThru", m_def_SelStrikeThru)
    URTB.SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
    URTB.SelRTF = PropBag.ReadProperty("SelRTF", m_def_SelRTF)
    URTB.SelRightIndent = PropBag.ReadProperty("SelRightIndent", m_def_SelRightIndent)
    URTB.SelProtected = PropBag.ReadProperty("SelProtected", m_def_SelProtected)
    URTB.SelLength = PropBag.ReadProperty("SelLength", m_def_SelLength)
    URTB.SelItalic = PropBag.ReadProperty("SelItalic", m_def_SelItalic)
    URTB.SelIndent = PropBag.ReadProperty("SelIndent", m_def_SelIndent)
    URTB.SelHangingIndent = PropBag.ReadProperty("SelHangingIndent", m_def_SelHangingIndent)
    URTB.SelFontSize = PropBag.ReadProperty("SelFontSize", m_def_SelFontSize)
    URTB.SelFontName = PropBag.ReadProperty("SelFontName", m_def_SelFontName)
    URTB.TextRTF = PropBag.ReadProperty("TextRTF", m_def_TextRTF)
    URTB.SelAlignment = PropBag.ReadProperty("SelAlignment", m_def_SelAlignment)
    URTB.SelBold = PropBag.ReadProperty("SelBold", m_def_SelBold)
    URTB.SelBullet = PropBag.ReadProperty("SelBullet", m_def_SelBullet)
    URTB.SelCharOffset = PropBag.ReadProperty("SelCharOffset", m_def_SelCharOffset)
    URTB.SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
    m_SelHColor = PropBag.ReadProperty("SelHColor", m_def_SelHColor)
    m_SelFontName = PropBag.ReadProperty("SelFontName", m_def_SelFontName)
End Sub

Private Sub UserControl_Resize()
    URTB.Move 0, 0, UserControl.Width - 60, UserControl.Height - 60
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", URTB.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", URTB.Enabled, True)
    Call PropBag.WriteProperty("Font", URTB.Font, Ambient.Font)
    Call PropBag.WriteProperty("AutoVerbMenu", URTB.AutoVerbMenu, False)
    Call PropBag.WriteProperty("BulletIndent", URTB.BulletIndent, 0)
    Call PropBag.WriteProperty("CausesValidation", URTB.CausesValidation, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Text", URTB.Text, "RichTextBox1")
    Call PropBag.WriteProperty("DisableNoScroll", URTB.DisableNoScroll, False)
    Call PropBag.WriteProperty("FileName", URTB.FileName, "")
    Call PropBag.WriteProperty("HideSelection", URTB.HideSelection, True)
    Call PropBag.WriteProperty("MaxLength", URTB.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", URTB.MousePointer, 0)
    Call PropBag.WriteProperty("RightMargin", URTB.RightMargin, 0)
    Call PropBag.WriteProperty("ScrollBars", URTB.ScrollBars, 0)
    Call PropBag.WriteProperty("WhatsThisHelpID", URTB.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("ToolTipText", URTB.ToolTipText, "")
    Call PropBag.WriteProperty("Locked", URTB.Locked, False)
    Call PropBag.WriteProperty("Transparent_RTF", m_Transparent_RTF, False)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("SelUnderline", URTB.SelUnderline, m_def_SelUnderline)
    Call PropBag.WriteProperty("SelText", URTB.SelText, m_def_SelText)
    Call PropBag.WriteProperty("SelTabCount", URTB.SelTabCount, m_def_SelTabCount)
    Call PropBag.WriteProperty("SelStrikeThru", URTB.SelStrikeThru, m_def_SelStrikeThru)
    Call PropBag.WriteProperty("SelStart", URTB.SelStart, m_def_SelStart)
    Call PropBag.WriteProperty("SelRTF", URTB.SelRTF, m_def_SelRTF)
    Call PropBag.WriteProperty("SelRightIndent", URTB.SelRightIndent, m_def_SelRightIndent)
    Call PropBag.WriteProperty("SelProtected", URTB.SelProtected, m_def_SelProtected)
    Call PropBag.WriteProperty("SelLength", URTB.SelLength, m_def_SelLength)
    Call PropBag.WriteProperty("SelItalic", URTB.SelItalic, m_def_SelItalic)
    Call PropBag.WriteProperty("SelIndent", URTB.SelIndent, m_def_SelIndent)
    Call PropBag.WriteProperty("SelHangingIndent", URTB.SelHangingIndent, m_def_SelHangingIndent)
    Call PropBag.WriteProperty("SelFontSize", URTB.SelFontSize, m_def_SelFontSize)
    Call PropBag.WriteProperty("SelFontName", URTB.SelFontName, m_def_SelFontName)
    Call PropBag.WriteProperty("TextRTF", URTB.TextRTF, m_def_TextRTF)
    Call PropBag.WriteProperty("SelAlignment", URTB.SelAlignment, m_def_SelAlignment)
    Call PropBag.WriteProperty("SelBold", URTB.SelBold, m_def_SelBold)
    Call PropBag.WriteProperty("SelBullet", URTB.SelBullet, m_def_SelBullet)
    Call PropBag.WriteProperty("SelCharOffset", URTB.SelCharOffset, m_def_SelCharOffset)
    Call PropBag.WriteProperty("SelColor", URTB.SelColor, m_def_SelColor)
    Call PropBag.WriteProperty("SelHColor", m_SelHColor, m_def_SelHColor)
    Call PropBag.WriteProperty("SelFontName", m_SelFontName, m_def_SelFontName)
End Sub

Public Property Get Appearance() As ggb
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
On Error Resume Next
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As ggb)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get BorderStyle() As gga
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
On Error Resume Next
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As gga)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get SelUnderline() As Boolean
Attribute SelUnderline.VB_Description = "Return or set font styles of the currently selected text in a RichTextBox control. The font styles include the following formats: Bold, Italic, Strikethru, and Underline. Not available at design time."
On Error Resume Next
    SelUnderline = URTB.SelUnderline
End Property

Public Property Let SelUnderline(ByVal New_SelUnderline As Boolean)
    URTB.SelUnderline = New_SelUnderline
    PropertyChanged "SelUnderline"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns or sets the string containing the currently selected text; consists of a zero-length string ("""") if no characters are selected."
On Error Resume Next
    SelText = URTB.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    URTB.SelText = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get SelTabCount() As Integer
Attribute SelTabCount.VB_Description = "Returns or sets the number of tabs and the absolute tab positions of text in a RichTextBox control. Not available at design time."
On Error Resume Next
    SelTabCount = URTB.SelTabCount
End Property

Public Property Let SelTabCount(ByVal New_SelTabCount As Integer)
    URTB.SelTabCount = New_SelTabCount
    PropertyChanged "SelTabCount"
End Property

Public Property Get SelStrikeThru() As Boolean
Attribute SelStrikeThru.VB_Description = "Return or set font styles of the currently selected text in a RichTextBox control. The font styles include the following formats: Bold, Italic, Strikethru, and Underline. Not available at design time."
On Error Resume Next
    SelStrikeThru = URTB.SelStrikeThru
End Property

Public Property Let SelStrikeThru(ByVal New_SelStrikeThru As Boolean)
    URTB.SelStrikeThru = New_SelStrikeThru
    PropertyChanged "SelStrikeThru"
End Property

Public Property Get SelStart() As Integer
Attribute SelStart.VB_Description = "Returns or sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
On Error Resume Next
    SelStart = URTB.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Integer)
    URTB.SelStart = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_Description = "Returns or sets the text (in .rtf format) in the current selection of a RichTextBox control. Not available at design time."
On Error Resume Next
    SelRTF = URTB.SelRTF
End Property

Public Property Let SelRTF(ByVal New_SelRTF As String)
    URTB.SelRTF = New_SelRTF
    PropertyChanged "SelRTF"
End Property

Public Property Get SelRightIndent() As Integer
Attribute SelRightIndent.VB_Description = "Returns or sets the margin settings for the paragraph(s) in a RichTextBox control that either contain the current selection or are added at the current insertion point. Not available at design time."
On Error Resume Next
    SelRightIndent = URTB.SelRightIndent
End Property

Public Property Let SelRightIndent(ByVal New_SelRightIndent As Integer)
    URTB.SelRightIndent = New_SelRightIndent
    PropertyChanged "SelRightIndent"
End Property

Public Property Get SelProtected() As Boolean
Attribute SelProtected.VB_Description = "Returns or sets a value which determines if the current selection is protected. Not available at design time."
On Error Resume Next
    SelProtected = URTB.SelProtected
End Property

Public Property Let SelProtected(ByVal New_SelProtected As Boolean)
    URTB.SelProtected = New_SelProtected
    PropertyChanged "SelProtected"
End Property

Public Property Get SelLength() As Integer
Attribute SelLength.VB_Description = "Returns or sets the number of characters selected.\r\n"
On Error Resume Next
    SelLength = URTB.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Integer)
    URTB.SelLength = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelItalic() As Boolean
Attribute SelItalic.VB_Description = "Returns or set font styles of the currently selected text in a RichTextBox control. The font styles include the following formats: Bold, Italic, Strikethru, and Underline. Not available at design time."
On Error Resume Next
    SelItalic = URTB.SelItalic
End Property

Public Property Let SelItalic(ByVal New_SelItalic As Boolean)
    URTB.SelItalic = New_SelItalic
    PropertyChanged "SelItalic"
End Property

Public Property Get SelIndent() As Integer
Attribute SelIndent.VB_Description = "Returns or sets the margin settings for the paragraph(s) in a RichTextBox control that either contain the current selection or are added at the current insertion point. Not available at design time."
On Error Resume Next
    SelIndent = URTB.SelIndent
End Property

Public Property Let SelIndent(ByVal New_SelIndent As Integer)
    URTB.SelIndent = New_SelIndent
    PropertyChanged "SelIndent"
End Property

Public Property Get SelHangingIndent() As Integer
Attribute SelHangingIndent.VB_Description = "Returns or sets the margin settings for the paragraph(s) in a RichTextBox control that either contain the current selection or are added at the current insertion point. Not available at design time."
On Error Resume Next
    SelHangingIndent = URTB.SelHangingIndent
End Property

Public Property Let SelHangingIndent(ByVal New_SelHangingIndent As Integer)
    URTB.SelHangingIndent = New_SelHangingIndent
    PropertyChanged "SelHangingIndent"
End Property

Public Property Get SelFontSize() As Integer
Attribute SelFontSize.VB_Description = "Returns or sets a value that specifies the size of the font used to display text in a RichTextBox control. Not available at design time."
On Error Resume Next
    SelFontSize = URTB.SelFontSize
End Property

Public Property Let SelFontSize(ByVal New_SelFontSize As Integer)
    URTB.SelFontSize = New_SelFontSize
    PropertyChanged "SelFontSize"
End Property
Public Property Get TextRTF() As String
Attribute TextRTF.VB_Description = "Returns or sets the text of a RichTextBox control, including all .rtf code."
On Error Resume Next
    TextRTF = URTB.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    URTB.TextRTF = New_TextRTF
    PropertyChanged "TextRTF"
End Property

Public Property Get SelAlignment() As AlignmentConstants
Attribute SelAlignment.VB_Description = "Returns or sets a value that controls the alignment of the paragraphs in a RichTextBox control. Not available at design time."
    SelAlignment = URTB.SelAlignment
End Property

Public Property Let SelAlignment(ByVal New_SelAlignment As AlignmentConstants)
    URTB.SelAlignment = New_SelAlignment
    PropertyChanged "SelAlignment"
End Property

Public Property Get SelBold() As Boolean
Attribute SelBold.VB_Description = "Returns or set font styles of the currently selected text in a RichTextBox control. The font styles include the following formats: Bold, Italic, Strikethru, and Underline. Not available at design time."
On Error Resume Next
    SelBold = URTB.SelBold
End Property

Public Property Let SelBold(ByVal New_SelBold As Boolean)
    URTB.SelBold = New_SelBold
    PropertyChanged "SelBold"
End Property

Public Property Get SelBullet() As Boolean
Attribute SelBullet.VB_Description = "Returns or sets a value that determines if a paragraph in the RichTextBox control containing the current selection or insertion point has the bullet style. Not available at design time."
    SelBullet = URTB.SelBullet
End Property

Public Property Let SelBullet(ByVal New_SelBullet As Boolean)
    URTB.SelBullet = New_SelBullet
    PropertyChanged "SelBullet"
End Property

Public Property Get SelCharOffset() As Boolean
Attribute SelCharOffset.VB_Description = "Returns or sets a value that determines whether text in the RichTextBox control appears on the baseline (normal), as a superscript above the baseline, or as a subscript below the baseline. Not available at design time."
    SelCharOffset = URTB.SelCharOffset
End Property

Public Property Let SelCharOffset(ByVal New_SelCharOffset As Boolean)
    URTB.SelCharOffset = New_SelCharOffset
    PropertyChanged "SelCharOffset"
End Property

Public Property Get SelColor() As OLE_COLOR
Attribute SelColor.VB_Description = "Returns or sets a value that determines the color of text in the RichTextBox control. Not available at design time."
On Error Resume Next
    SelColor = URTB.SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As OLE_COLOR)
    URTB.SelColor = New_SelColor
    PropertyChanged "SelColor"
End Property

Private Sub UserControl11_GotFocus()

End Sub

Public Property Get SelHColor() As OLE_COLOR
Attribute SelHColor.VB_Description = "Returns or sets a value that determines the highlight color of text in the RichTextBox control. Not available at design time."
On Error Resume Next
    SelHColor = m_SelHColor
End Property

Public Property Let SelHColor(ByVal New_SelHColor As OLE_COLOR)
    m_SelHColor = New_SelHColor
        udtCharFormat.dwMask = CFM_BACKCOLOR
        udtCharFormat.cbSize = LenB(udtCharFormat)
        udtCharFormat.crBackColor = m_SelHColor
    Call SendMessageByVal(URTB.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(udtCharFormat))
    PropertyChanged "SelHColor"
End Property

Public Property Get SelFontName() As String
Attribute SelFontName.VB_Description = "Returns or sets the font used to display the currently selected text or the character(s) immediately following the insertion point in the RichTextBox control. Not available at design time."
On Error Resume Next
    SelFontName = URTB.SelFontName
End Property

Public Property Let SelFontName(ByVal New_SelFontName As String)
    URTB.SelFontName = New_SelFontName
    PropertyChanged "SelFontName"
End Property

