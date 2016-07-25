VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRawRequest 
   Caption         =   "Send Raw HTTP Request"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   -105
      TabIndex        =   1
      Top             =   2625
      Width           =   6105
      Begin VB.CommandButton cmdRawRequest 
         Caption         =   "Send Raw Request"
         Height          =   435
         Left            =   4200
         TabIndex        =   4
         Top             =   0
         Width           =   1800
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   0
         Width           =   2850
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         TabIndex        =   2
         Text            =   "80"
         Top             =   0
         Width           =   435
      End
      Begin MSWinsockLib.Winsock ws 
         Left            =   3675
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2430
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   105
      Width           =   5790
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuViewWebPage 
         Caption         =   "View Page in IE"
      End
      Begin VB.Menu mnuParseHTML 
         Caption         =   "Remove HTML"
      End
      Begin VB.Menu mnuReqHeaders 
         Caption         =   "Back to Req Headers"
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEscape 
         Caption         =   "URL Decode"
         Index           =   0
      End
      Begin VB.Menu mnuEscape 
         Caption         =   "URL Encode"
         Index           =   1
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Base64 Decode Sel"
         Index           =   0
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Base64 Encode Sel"
         Index           =   1
      End
      Begin VB.Menu mnuspacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlugins 
         Caption         =   "Plugins"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRawRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the raw request functions will probably fail on chunked transfers

Private GeneratedHeaders As String


Private Sub cmdRawRequest_Click()
    On Error GoTo oops
    ws.Close
    If Text1 = Empty Or Text2 = Empty Or Text3 = Empty Then MsgBox "Not enough input fill in textboxes": Exit Sub
    If InStr(Text1, vbCrLf & vbCrLf) < 1 Then Text1 = Text1 & vbCrLf & vbCrLf
    GeneratedHeaders = Text1
    ws.Connect Text2, Text3
    Exit Sub
oops: MsgBox Err.Description, vbExclamation
End Sub

Sub PromptForUrlThenInitalize()
    ret = InputBox("Enter the url you wish to request", , frmMain.txtUrl)
    If Left(ret, 10) = Left(frmMain.txtUrl, 10) Then useCookie = True
    PrepareRawRequest ret, IIf(useCookie, frmMain.wb.Document.Cookie, Empty)
End Sub

Sub PrepareRawRequest(url, Optional Cookie = Empty)
    url = Trim(LTrim(Replace(url, "http://", "")))
    
    slash = InStr(url, "/")
    If slash > 2 Then
        Text2 = Mid(url, 1, slash - 1)
        url = Mid(url, slash + 1, Len(url))
    Else
        Text2 = url
        url = "/"
    End If
    
    Cookie = IIf(Cookie <> Empty, "Cookie: " & Cookie & "\n", Empty)
    Text1 = "GET /" & url & br(" HTTP/1.1\nHOST: " & Text2.text & "\nReferer: " & Text2.text & "\nACCEPT: */*\nAccept-Encoding: None\nUser-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0)\nConnection: Close\nAccept-Transfer-Encoding: None\n" & Cookie & "\n")
    Me.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
   Select Case KeyAscii
        Case 1: Text1.SelStart = 0: Text1.SelLength = Len(Text1) 'ctrl-A =selall
        Case 3, 24: 'ctrl-C = copy
                If Text1.SelLength = 0 Then Exit Sub
                Clipboard.Clear: Clipboard.SetText Text1.SelText
        Case 24: 'ctrl-X = cut
                Text1.SelText = Empty
        Case 19 'ctrl-S = insert char x times
                c = InputBox("Enter Character you wish to insert x times at current cursor location the default example inserts A 256 times", , "A 256")
                If c = Empty Or InStr(c, " ") < 1 Then Exit Sub
                tmp = Split(c, " ")
                Text1.SelText = String(CLng(tmp(1)), tmp(0))
        Case 4 'ctrl-D = count chars in selected string
                If Text1.SelText = Empty Then Exit Sub
                MsgBox "SelChar Count: " & Len(Text1.SelText)
  End Select
End Sub

Private Sub ws_Connect()
    ws.SendData Text1
    Text1 = Empty
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim tmp As String
    ws.GetData tmp, vbString
    If Len(Text1) < 50000 Then Text1 = Text1 & tmp
End Sub

Private Sub ws_Close()
    ws.Close
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then ShowRtClkMenu Me, Text1, mnuPopup
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Height = Me.Height - 1150
    Text1.Width = Me.Width - 300
    Frame.Top = Text1.Height + Text1.Top + 150
    Frame.Left = Me.Width - Frame.Width - 200
End Sub

Private Sub Form_Load()
    n = "frmRawHttp"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
End Sub

Private Sub mnuViewWebPage_Click()
   frmMain.RenderPage Text1
End Sub

Private Sub mnuReqHeaders_Click()
    Text1 = GeneratedHeaders
End Sub

Private Sub mnuPlugins_Click(index As Integer)
    FirePluginEvent index, "frmRawRequest.mnuPlugins"
End Sub

Private Sub mnuParseHTML_Click()
  If Text1.SelText = Empty Then Exit Sub
  Text1.SelText = parseHtml(Text1.SelText)
End Sub

Private Sub mnuDecode_Click(index As Integer)
    If Text1.SelText = Empty Then Exit Sub
    If index = 1 Then Text1.SelText = b64Encode(Text1.SelText) _
    Else: Text1.SelText = b64Decode(Text1.SelText)
End Sub

Private Sub mnuEscape_Click(index As Integer)
    If Text1.SelText = Empty Then Exit Sub
    If index = 1 Then Text1.SelText = escape(Text1.SelText) _
    Else Text1.SelText = UnEscape(Text1.SelText)
End Sub

Function br(it)
    br = Replace(it, "\n", vbCrLf)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'if the frmUnloads completly when they close it with the x then
    'any additions to the rt click menu from plugins will disappear!
    Cancel = 1
    Me.Hide
    
    n = "frmRawHttp"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
    End If
End Sub



'screw it, its easier to make sure url is absolute by eye
'make sure they are absolute URL's
'Function MakeURLAbsolute(url, pageUrl) As String
'
'    If InStr(url, "http://") > 0 Then MakeURLAbsolute = url: Exit Function
'
'    pageUrl = Replace(pageUrl, "http://", Empty)
'    s = InStr(pageUrl, "/")
'    If s > 0 Then
'        host = Mid(pageUrl, 1, s - 1)
'        folder = Mid(pageUrl, s, Len(pageUrl))
'    Else
'        host = pageUrl
'        folder = Empty
'    End If
'
'    s = InStr(url, "/")
'    If s < 1 Then
'        MakeURLAbsolute = host & folder & url
'    Else
'        urlFolder = split(
'
'
'End Function
