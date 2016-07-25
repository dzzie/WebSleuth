VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRawRequest 
   Caption         =   "Send Raw HTTP Request"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox chkWordWrap 
      Caption         =   "WordWrap"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   3480
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   6480
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   6480
      Top             =   0
   End
   Begin TabDlg.SSTab objTab 
      Height          =   4035
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7117
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      TabCaption(0)   =   "Request"
      TabPicture(0)   =   "frmRawRequest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstQueryString"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "rtfReq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkShowQSInline"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Response"
      TabPicture(1)   =   "frmRawRequest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameRTF"
      Tab(1).Control(1)=   "rtfResponse"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Render"
      TabPicture(2)   =   "frmRawRequest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "wbRender"
      Tab(2).ControlCount=   1
      Begin VB.CheckBox chkShowQSInline 
         Caption         =   "Show Args Inline"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Frame frameRTF 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   -74820
         TabIndex        =   9
         Top             =   3060
         Width           =   6375
         Begin VB.ComboBox cboColor 
            Height          =   315
            ItemData        =   "frmRawRequest.frx":0054
            Left            =   2640
            List            =   "frmRawRequest.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   60
            Width           =   1575
         End
         Begin VB.CommandButton cmdHighLight 
            Caption         =   "Highlight All"
            Height          =   315
            Left            =   4320
            TabIndex        =   16
            Top             =   60
            Width           =   1575
         End
         Begin VB.TextBox txtFind 
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   60
            Width           =   1455
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   315
            Left            =   1560
            TabIndex        =   10
            Top             =   60
            Width           =   795
         End
         Begin VB.Image ImgSave 
            Height          =   240
            Left            =   6060
            Picture         =   "frmRawRequest.frx":0058
            Top             =   120
            Width           =   240
         End
      End
      Begin WebSleuth.RTF rtfReq 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4048
      End
      Begin WebSleuth.List lstQueryString 
         Height          =   990
         Left            =   120
         TabIndex        =   6
         Top             =   2460
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   1746
      End
      Begin WebSleuth.RTF rtfResponse 
         Height          =   2835
         Left            =   -74880
         TabIndex        =   7
         Top             =   120
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   5001
      End
      Begin SHDocVwCtl.WebBrowser wbRender 
         Height          =   3435
         Left            =   -74880
         TabIndex        =   8
         Top             =   60
         Width           =   6435
         ExtentX         =   11351
         ExtentY         =   6059
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   4140
      Width           =   6765
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close Sck"
         Height          =   375
         Left            =   5700
         TabIndex        =   14
         Top             =   0
         Width           =   915
      End
      Begin VB.CheckBox chkPost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   12
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRawRequest 
         Caption         =   "Send Req"
         Height          =   375
         Left            =   4620
         TabIndex        =   3
         Top             =   0
         Width           =   960
      End
      Begin VB.TextBox txtHost 
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
         Left            =   180
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   0
         Width           =   2850
      End
      Begin VB.TextBox txtPort 
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
         TabIndex        =   1
         Text            =   "80"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuParseHTML 
         Caption         =   "Remove HTML"
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
Private LastUrl As String
Private LastCookie As String

Private Sub chkPost_Click()
    If chkPost.Tag = Empty Then
        PrepareRawRequest LastUrl, LastCookie, CBool(chkPost.value)
    Else
        'so we can change the chkPost.Value withouth triggering above
        chkPost.Tag = Empty
    End If
End Sub

Private Sub chkShowQSInline_Click()
    lstQueryString.Visible = IIf(chkShowQSInline.value = 1, False, True)
    PrepareRawRequest LastUrl, LastCookie, CBool(chkPost.value)
    Form_Resize
End Sub

Private Sub chkWordWrap_Click()
    If objTab.Tab = 0 Then
        rtfReq.WordWrap = IIf(chkWordWrap.value = 1, True, False)
    Else
        rtfResponse.WordWrap = IIf(chkWordWrap.value = 1, True, False)
    End If
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    ws.Close
    Me.caption = "Socket Closed"
End Sub

Private Sub cmdFind_Click()
   If txtFind <> rtfResponse.FindString Then
        rtfResponse.FindString = txtFind
        rtfResponse.find
    Else
        rtfResponse.findNext
    End If
End Sub

Private Sub cmdRawRequest_Click()
    On Error GoTo oops
    ws.Close
    If txtHost = Empty Or txtPort = Empty Or txtPort = Empty Then MsgBox "Not enough input fill in textboxes": Exit Sub
    If lstQueryString.Count > 0 And InStr(rtfReq.text, "<QUERYSTRING>") < 1 Then
        MsgBox "Ughh <QUERYSTRING> marker not found but list arguments exist. That tag will be replaced with the actual querystring args from the list in the request just so you know."
    End If
    If InStr(rtfReq.text, vbCrLf & vbCrLf) < 1 Then rtfReq.AppendIt vbCrLf & vbCrLf
    GeneratedHeaders = rtfReq.text
    Me.caption = "Attempting Connect to :" & txtHost & " Port:" & txtPort
    tmrTimeout.Enabled = True
    ws.Connect txtHost, txtPort
    Exit Sub
oops: MsgBox Err.Description, vbExclamation
End Sub

Sub PromptForUrlThenInitalize(Optional ByVal usePOST = False)
    ret = frmAnalyze.AnlyzeUrl(frmMain.txtUrl, , False, "Enter the url you wish to request")
    If ret = Empty Then Exit Sub
    If Left(ret, 10) = Left(frmMain.txtUrl, 10) Then useCookie = True
    PrepareRawRequest ret, IIf(useCookie, frmMain.wb.Document.cookie, Empty), usePOST
End Sub

Sub PrepareRawRequest(ByVal url, Optional ByVal cookie = Empty, Optional ByVal usePOST = False)
    
    LastUrl = url
    LastCookie = Replace(cookie, "Cookie: ", Empty)
    
    If usePOST And chkPost.value <> 1 Then chkPost.Tag = "dont trigger": chkPost.value = 1
    
    If InStr(url, "http://") < 1 Then url = "http://" & url
    lstQueryString.Clear
    
    url = Replace(url, "https://", "http://")
    url = Trim(LTrim(Replace(url, "http://", "")))
    
    slash = InStr(url, "/")
    ques = InStr(url, "?")
    
    If slash > 2 Then
        Call SetHostPort(url, slash)
        If ques > 0 And chkShowQSInline.value = 0 Then
            If ques < slash Then MsgBox "awe booger: ques < slash"
            path = Mid(url, slash + 1, ques - slash) & "<QUERYSTRING>"
            qs = Mid(url, ques + 1, Len(url))
            lstQueryString.LoadDelimitedString qs, "&"
        Else
            path = Mid(url, slash + 1, Len(url))
        End If
    Else
        Call SetHostPort(url, Len(url) + 1)
    End If
    
    cookie = IIf(cookie <> Empty, "Cookie: " & cookie & "\n", Empty)
    If Not usePOST Then
        rtfReq.text = "GET /" & path & br(" HTTP/1.1\nHOST: " & txtHost.text & "\nReferer: " & txtHost.text & "\nACCEPT: */*\nAccept-Encoding: None\nUser-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0)\nConnection: Close\nAccept-Transfer-Encoding: None\n" & cookie & "\n")
    Else
        rtfReq.text = "POST /" & Replace(path, "?<QUERYSTRING>", Empty) & br(" HTTP/1.1\nHOST: " & txtHost.text & "\nReferer: " & txtHost.text & "\nACCEPT: */*\nAccept-Encoding: None\nUser-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0)\nConnection: Close\nAccept-Transfer-Encoding: None\nContent-Type: application/x-www-form-urlencoded\nContent-Length: " & Len(lstQueryString.GetListContents("&")) & "\n" & cookie & "\n<QUERYSTRING>")
    End If
    
    Me.Show
End Sub

Sub SetHostPort(ByVal url, ByVal slash)
        If url = "about:blank" Then
            'this is inevitable so deal i guess
            txtHost = url: txtPort = 80: Exit Sub
        End If
        
        hst = Mid(url, 1, slash - 1)
        semiC = InStr(hst, ":")
        If semiC > 0 Then
            txtHost = Mid(hst, 1, semiC - 1)
            txtPort = Mid(hst, semiC + 1, Len(hst))
        Else
            txtHost = hst
            txtPort = 80
        End If
End Sub


Private Sub ImgSave_Click()
    On Error Resume Next
    ans = frmMain.CmnDlg1.ShowSave(App.path, textFiles, "Save Source As:")
    If ans = Empty Then Exit Sub
    WriteFile ans, rtfResponse.text
End Sub

Private Sub lstQueryString_DoubleClick()
    With lstQueryString
        oldValue = .SelectedText
        newValue = frmAnalyze.AnlyzeUrl(oldValue, , False)
        .SelectedText = newValue
        If newValue <> oldValue And chkPost.value = 1 Then
            'we need to update content length so gen new header
            PrepareRawRequest Replace(LastUrl, oldValue, newValue), LastCookie, True
        End If
    End With
End Sub

Private Sub objTab_Click(PreviousTab As Integer)
    chkWordWrap.Visible = IIf(objTab.Tab < 2, True, False)
    Select Case objTab.Tab
        Case 0: chkWordWrap.value = IIf(rtfReq.WordWrap, 1, 0)
        Case 1: chkWordWrap.value = IIf(rtfResponse.WordWrap, 1, 0)
    End Select
    
    If PreviousTab = 1 And objTab.Tab = 2 Then
        If Not setWbBody(rtfResponse.text) Then
              wbRender.Document.write rtfResponse.text
        End If
    End If
End Sub

Private Function setWbBody(html) As Boolean
 On Error GoTo oops
    wbRender.Document.body.innerHTML = html
    setWbBody = True
 Exit Function
oops:
    setWbBody = False
End Function

Private Sub rtfReq_KeyPress(KeyAscii As Integer)
On Error Resume Next
   Select Case KeyAscii
        Case 1: rtfReq.SelectAll  'ctrl-A =selall
        Case 3: rtfReq.CopySelection  'ctrl-C = copy
        Case 24: rtfReq.Cut 'ctrl-X = cut
        Case 19 'ctrl-S = insert char x times
                c = InputBox("Enter Character you wish to insert x times at current cursor location the default example inserts A 256 times", , "A 256")
                If c = Empty Or InStr(c, " ") < 1 Then Exit Sub
                tmp = Split(c, " ")
                rtfReq.SelText = String(CLng(tmp(1)), tmp(0))
        Case 4 'ctrl-D = count chars in selected string
                If rtfReq.SelText = Empty Then Exit Sub
                MsgBox "SelChar Count: " & rtfReq.SelLen
  End Select
End Sub

Private Sub rtfReq_RightClick()
    
    ShowRtClkMenu Me, rtfReq, mnuPopup
End Sub

Private Sub rtfResponse_RightClick()
    ShowRtClkMenu Me, rtfReq, mnuPopup
End Sub

Private Sub tmrTimeout_Timer()
    tmrTimeout.Enabled = False
    If MsgBox(br("We have not connected withing the 4 second timeout period.\n\nDo you want to continue waiting?\n\nNote: This is Winsock only. If you need to access a proxy to hit the net you will have to set it up manually in the request!."), vbExclamation + vbYesNo) = vbYes Then
        tmrTimeout.Enabled = True
    Else
        Me.caption = "Request Timedout..."
        cmdClose_Click
    End If
End Sub

Private Sub ws_Connect()
    tmrTimeout.Enabled = False
    Me.caption = "Host found Connecting..."
    ws.SendData Replace(rtfReq.text, "<QUERYSTRING>", lstQueryString.GetListContents("&"))
    rtfResponse.text = Empty
    wbRender.Navigate2 "about:blank"
    objTab.Tab = 1
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim tmp As String
    Me.caption = "Receiving Data..."
    ws.GetData tmp, vbString
    rtfResponse.AppendIt tmp
End Sub

Private Sub ws_Close()
    ws.Close
    Me.caption = "Connection Closed..."
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height > 4000 Then
        objTab.Height = Me.Height - Frame.Height - 600
        wbRender.Height = objTab.Height - 600
        rtfResponse.Height = objTab.Height - frameRTF.Height - 800
        frameRTF.Top = rtfResponse.Height + 200
        Frame.Top = objTab.Height + objTab.Top + 150
        If chkShowQSInline.value = 0 Then
            rtfReq.Height = objTab.Height - lstQueryString.Height - 800
        Else
            rtfReq.Height = objTab.Height - 800
        End If
        lstQueryString.Top = rtfReq.Height + rtfReq.Top + 100
    End If
    If Me.Width > 4000 Then
        objTab.Width = Me.Width - 150
        rtfReq.Width = objTab.Width - 150
        rtfResponse.Width = rtfReq.Width
        lstQueryString.Width = rtfReq.Width
        wbRender.Width = rtfReq.Width
        chkShowQSInline.Top = objTab.Height - 550
        chkWordWrap.Top = chkShowQSInline.Top
        Frame.Left = Me.Width - Frame.Width - 200
    End If
End Sub

Private Sub Form_Load()
    n = "frmRawHttp"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
    chkShowQSInline.value = CInt(GetSetting(App.title, n, "ShowQS", 0))
    
    wbRender.Navigate2 "about:blank"
    objTab.Tab = 0
    chkWordWrap.value = IIf(rtfReq.WordWrap, 1, 0)
    
    c = Split("Red,Blue,Somthing,Green", ",")
    For i = 0 To UBound(c)
        cboColor.AddItem c(i)
    Next
    cboColor.ListIndex = 0
    
End Sub

Private Sub cmdHighlight_Click()
    LockWindowUpdate rtfResponse.hWnd
    c = Array(vbRed, vbBlue, &HC000C0, &H808000)
    rtfResponse.SetColor txtFind, CLng(c(cboColor.ListIndex)), , True
    rtfResponse.ScrollToTop
    LockWindowUpdate 0&
End Sub


Private Sub mnuPlugins_Click(index As Integer)
    FirePluginEvent index, "frmRawRequest.mnuPlugins"
End Sub

Private Sub mnuParseHTML_Click()
  If objTab.TabIndex = 1 Then
    If rtfResponse.SelText = Empty Then Exit Sub
    rtfResponse.SelText = parseHtml(rtfResponse.SelText)
  End If
End Sub

Private Sub mnuDecode_Click(index As Integer)
    Dim objRtfCtrl As rtf
    Set objRtfCtrl = IIf(objTab.TabIndex = 0, rtfReq, rtfResponse)
    If objRtfCtrl.SelText = Empty Then Exit Sub
    If index = 1 Then
        objRtfCtrl.SelText = b64Encode(objRtfCtrl.SelText)
    Else
        objRtfCtrl.SelText = b64Decode(objRtfCtrl.SelText)
    End If
    Set objRtfCtrl = Nothing
End Sub

Private Sub mnuEscape_Click(index As Integer)
    Dim objRtfCtrl As rtf
    Set objRtfCtrl = IIf(objTab.TabIndex = 0, rtfReq, rtfResponse)
    If objRtfCtrl.SelText = Empty Then Exit Sub
    If index = 1 Then objRtfCtrl.SelText = escape(objRtfCtrl.SelText) _
    Else objRtfCtrl.SelText = UnEscape(objRtfCtrl.SelText)
    Set objRtfCtrl = Nothing
End Sub

Function br(it)
    br = Replace(it, "\n", vbCrLf)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'if the frmUnloads completly when they close it with the x then
    'any additions to the rt click menu from plugins will disappear!
    Cancel = 1
    Me.Hide
    
    rtfResponse.text = Empty
    rtfReq.text = Empty
    lstQueryString.Clear
    wbRender.Navigate2 "about:blank"
    tmrTimeout.Enabled = False
    objTab.Tab = 0
    
    n = "frmRawHttp"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
        SaveSetting App.title, n, "ShowQS", chkShowQSInline.value
    End If
End Sub
