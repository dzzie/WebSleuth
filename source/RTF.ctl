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
Private Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const MM_TWIPS = 6
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9

Private startAt As Long
Private findWhat As String
Private ww As Boolean

Public Event RightClick()

Public Property Let text(it)
    rtb.text = it
End Property

Public Property Get text()
    text = rtb.text
End Property

Public Property Get FindString()
    FindString = findWhat
End Property

Public Property Let FindString(it)
    findWhat = CStr(it)
End Property

Public Property Let WordWrap(on_ As Boolean)
    ww = on_
    If on_ Then rtb.rightMargin = 0 Else Call SetRightMargain
End Property

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    rtb.text = Replace(rtb.text, find, changeTo, , , vbTextCompare)
End Sub

Sub SetColor(str, Optional color As ColorConstants = vbBlack, Optional fSize = 10, Optional bold As Boolean = False, Optional italic As Boolean = False)
    Y = 1
    X = rtb.find(str, Y)
    While X > 0
        rtb.SelStart = X
        rtb.SelLength = Len(str)
        rtb.SelColor = color
        rtb.SelBold = bold
        rtb.SelItalic = italic
        rtb.SelFontSize = fSize
        Y = X + 1
        X = rtb.find(str, Y)
    Wend
End Sub

Sub SetSpanColor(startStr, endStr, Optional color = &HFFFFFF, Optional fSize = 10, Optional bold As Boolean = False, Optional italic As Boolean = False)
    On Error Resume Next
    Y = 1
    X = rtb.find(startStr, Y)
    z = rtb.find(endStr, Y)
    While X > 0 And z > 0
        rtb.SelStart = X
        rtb.SelLength = z - X + Len(endStr)
        rtb.SelColor = color
        rtb.SelBold = bold
        rtb.SelItalic = italic
        rtb.SelFontSize = fSize
        Y = z + 1 + Len(endStr)
        X = rtb.find(startStr, Y)
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
    findWhat = InputBox("Find:", , findWhat)
    If findWhat = Empty Then Exit Sub
    Me.ScrollToTop
    startAt = 1
    X = rtb.find(findWhat)
    If X > 0 Then
        rtb.SelStart = X
        rtb.SelLength = Len(findWhat)
        startAt = X + 1
    Else
        MsgBox "String Not Found", vbInformation
    End If
End Sub

Sub findNext()
    If findWhat = Empty Then Exit Sub
    X = rtb.find(findWhat, startAt)
    If X > 0 And startAt < (Len(rtb.text) - 1) Then
        rtb.SelStart = X
        rtb.SelLength = Len(findWhat)
        startAt = X + 1
    Else
        startAt = 1
        MsgBox "Search Complete", vbInformation
    End If
End Sub


Private Sub rtb_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 6: find
        Case 4: findNext
        Case Else: 'msgBox KeyAscii
    End Select
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


