VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl RTF 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ScaleHeight     =   1410
   ScaleWidth      =   4170
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1170
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   2064
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"RTF.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "RTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type


Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const MM_TWIPS = 6
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9

Private startAt As Long
Private findWhat As String
Private ww As Boolean
Private dirty As Boolean

Public Event RightClick()
Public Event KeyPress(keyCode As Integer)

Public Property Let text(it)
    rtb.text = it
    dirty = False
End Property

Public Property Get text()
    text = rtb.text
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = dirty
End Property
Public Property Let IsDirty(x As Boolean)
    dirty = x
End Property

Public Property Get FindString()
    FindString = CollapseConstants(findWhat)
End Property

Public Property Let FindString(it)
    findWhat = ExpandConstants(it)
End Property

Function ExpandConstants(ByVal strIn) As String
    strIn = Replace(strIn, "<TAB>", vbTab, , , vbTextCompare)
    strIn = Replace(strIn, "<CRLF>", vbCrLf, , , vbTextCompare)
    strIn = Replace(strIn, "<CR>", vbCr, , , vbTextCompare)
    ExpandConstants = CStr(Replace(strIn, "<LF>", vbLf, , , vbTextCompare))
End Function

Function CollapseConstants(ByVal strIn) As String
    strIn = Replace(strIn, vbTab, "<TAB>", , , vbTextCompare)
    strIn = Replace(strIn, vbCrLf, "<CRLF>", , , vbTextCompare)
    strIn = Replace(strIn, vbCr, "<CR>", , , vbTextCompare)
    CollapseConstants = CStr(Replace(strIn, vbLf, "<LF>", , , vbTextCompare))
End Function

Public Property Let WordWrap(on_ As Boolean)
    ww = on_
    If on_ Then rtb.rightMargin = 0 Else Call SetRightMargain
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = ww
End Property

Public Property Let SelText(txt)
    rtb.SelText = txt
End Property

Public Property Get SelText()
    SelText = rtb.SelText
End Property

Public Property Let SelStart(x)
    rtb.SelStart = x
End Property

Public Property Get SelStart()
    SelStart = rtb.SelStart
End Property

Public Property Let SelLen(x)
    rtb.SelLength = x
End Property

Public Property Get SelLen()
    SelLen = rtb.SelLength
End Property

Public Property Let Enabled(t As Boolean)
    rtb.Enabled = t
End Property

Public Property Get Enabled() As Boolean
    Enabled = rtb.Enabled
End Property

Public Property Get hWnd() As Long
    hWnd = rtb.hWnd
End Property

Sub SelSpan(start, length)
    On Error Resume Next
    rtb.SelStart = start
    rtb.SelLength = length
End Sub

Sub CopySelection()
    If rtb.SelLength < 1 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText rtb.SelText
End Sub

Sub Cut()
    CopySelection
    rtb.SelText = Empty
End Sub

Sub SelectAll()
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.text)
End Sub

Private Sub rtb_Change()
    dirty = True
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then RaiseEvent RightClick
End Sub

Private Sub UserControl_Initialize()
    rtb.Left = 0
    rtb.Top = 0
    startAt = 1
End Sub

Sub ScrollToTop()
    rtb.SelStart = 1
    rtb.SelLength = 0
    On Error Resume Next
    rtb.SetFocus
End Sub

Sub ReplaceText(find, changeTo)
    rtb.text = Replace(rtb.text, ExpandConstants(find), ExpandConstants(changeTo), , , vbTextCompare)
End Sub

Sub SetColor(str, Optional color As ColorConstants = vbBlack, Optional fSize = 10, Optional bold As Boolean = False, Optional italic As Boolean = False)
    Y = 1
    x = rtb.find(str, Y)
    While x > 0
        rtb.SelStart = x
        rtb.SelLength = Len(str)
        rtb.SelColor = color
        rtb.SelBold = bold
        rtb.SelItalic = italic
        rtb.SelFontSize = fSize
        Y = x + 1
        x = rtb.find(str, Y)
    Wend
End Sub

Sub SetSpanColor(startStr, endStr, Optional color = &HFFFFFF, Optional fSize = 10, Optional bold As Boolean = False, Optional italic As Boolean = False)
    On Error Resume Next
    Y = 1
    x = rtb.find(startStr, Y)
    z = rtb.find(endStr, Y)
    While x > 0 And z > 0
        rtb.SelStart = x
        rtb.SelLength = z - x + Len(endStr)
        rtb.SelColor = color
        rtb.SelBold = bold
        rtb.SelItalic = italic
        rtb.SelFontSize = fSize
        Y = z + 1 + Len(endStr)
        x = rtb.find(startStr, Y)
        z = rtb.find(endStr, Y)
    Wend
End Sub
Private Sub UserControl_Resize()
        rtb.Height = UserControl.Height
        rtb.Width = UserControl.Width
End Sub

Sub MatchSize(it As Object)
    UserControl.Size it.Width, it.Height
End Sub

Sub find()
    'findWhat = InputBox("Find:", , findWhat)
    If findWhat = Empty Then Exit Sub
    Me.ScrollToTop
    startAt = 1
    x = rtb.find(findWhat)
    If x >= 0 Then
        rtb.SelStart = x
        rtb.SelLength = Len(findWhat)
        startAt = x + 1
    Else
        MsgBox "String Not Found", vbInformation
    End If
End Sub

Sub findNext()
    If findWhat = Empty Then Exit Sub
    x = rtb.find(findWhat, startAt)
    If x > 0 And startAt < (Len(rtb.text) - 1) Then
        rtb.SelStart = x
        rtb.SelLength = Len(findWhat)
        startAt = x + 1
    Else
        startAt = 1
        MsgBox "Search Complete", vbInformation
    End If
End Sub


Private Sub rtb_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 6: find
        Case 4: findNext
        Case Else: RaiseEvent KeyPress(KeyAscii)
    End Select
End Sub

Sub AppendIt(it)
    rtb.text = rtb.text & it
End Sub

Sub PrePendIt(it)
    rtb.text = it & rtb.text
End Sub

Sub SetRightMargain()
    Dim tm As TEXTMETRIC
    Dim ret&, hWnd As Long, hdc As Long
    Dim lnglength&
    Dim currLine&, lineCount&, lineLength&, lineIndex&
    
    lineCount = SendMessageLong(rtb.hWnd, EM_GETLINECOUNT, 0&, 0&)
    
    For i = 0 To lineCount - 1
        lineIndex = SendMessageLong(rtb.hWnd, EM_LINEINDEX, i, 0&)
        lineLength = SendMessageLong(rtb.hWnd, EM_LINELENGTH, lineIndex, 0&)
        If lineLength > ret Then ret = lineLength
    Next
        
    hWnd = rtb.hWnd
    hdc = GetWindowDC(hWnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC hWnd, hdc
    End If
    
    rtb.rightMargin = ret * tm.tmMaxCharWidth
End Sub

Sub highlightHtml()
    SetSpanColor "<form", ">", vbRed, , True
    SetColor "</form>", vbRed, , True
    SetColor "<input", vbBlue, , True
    SetColor "<script", &HC000C0, , True
    SetColor "</script>", &HC000C0, , True
    SetSpanColor "<!--", "-->", &H808000
    ScrollToTop
End Sub
