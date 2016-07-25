VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPauseNavigate 
      Enabled         =   0   'False
      Interval        =   1700
      Left            =   8520
      Top             =   3840
   End
   Begin VB.Frame splitter 
      BackColor       =   &H00808080&
      Height          =   60
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   16
      Top             =   2835
      Width           =   9435
   End
   Begin WebSleuth.CmnDlg CmnDlg1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   582
      _ExtentY        =   503
   End
   Begin WebSleuth.RTF rtf 
      Height          =   1800
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3175
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2760
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9435
      ExtentX         =   16642
      ExtentY         =   4868
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
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   0
      TabIndex        =   0
      Top             =   2730
      Width           =   9465
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   800
         Left            =   5040
         TabIndex        =   8
         Top             =   210
         Width           =   4335
         Begin VB.CommandButton cmdNavigate 
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   2
            Left            =   945
            TabIndex        =   13
            ToolTipText     =   "Stop Navigation"
            Top             =   0
            Width           =   930
         End
         Begin VB.CommandButton cmdNavigate 
            Caption         =   "Frwd >>"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   1
            Left            =   1875
            TabIndex        =   12
            ToolTipText     =   "Navigate forward"
            Top             =   0
            Width           =   960
         End
         Begin VB.CommandButton cmdNavigate 
            Caption         =   "<< Back"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   0
            Left            =   0
            TabIndex        =   11
            ToolTipText     =   "Navigate back"
            Top             =   0
            Width           =   930
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Source"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3045
            TabIndex        =   10
            ToolTipText     =   "edit the source code of the remote webpage"
            Top             =   0
            Width           =   1245
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2520
            TabIndex        =   9
            Text            =   "*"
            Top             =   420
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apply Result Filter"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   435
            Width           =   1785
         End
         Begin VB.Label lblFilter 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LIKE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1890
            TabIndex        =   14
            ToolTipText     =   "click to toggle"
            Top             =   435
            Width           =   585
         End
      End
      Begin WebSleuth.List List 
         Height          =   1470
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "list a range of properties selected with combo box..rt click for tools"
         Top             =   1050
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2593
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   525
         Width           =   3135
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Urls"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   2640
            TabIndex        =   19
            ToolTipText     =   "oghh oghh right click me!"
            Top             =   135
            Width           =   360
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2520
            X2              =   2520
            Y1              =   105
            Y2              =   420
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   1740
            X2              =   1740
            Y1              =   105
            Y2              =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Plugins"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   1800
            TabIndex        =   18
            ToolTipText     =   "oghh oghh right click me!"
            Top             =   135
            Width           =   675
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   890
            X2              =   890
            Y1              =   105
            Y2              =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   945
            TabIndex        =   17
            ToolTipText     =   "oghh oghh right click me!"
            Top             =   135
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Actions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   95
            TabIndex        =   6
            ToolTipText     =   "oghh oghh right click me!"
            Top             =   135
            Width           =   735
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmMain.frx":030A
         Left            =   105
         List            =   "frmMain.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "View a specific attribute of the web page"
         Top             =   600
         Width           =   1590
      End
      Begin VB.TextBox txtUrl 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Addressbar...enter url to surf to and hit return"
         Top             =   210
         Width           =   4785
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "mnuAdvanced"
      Visible         =   0   'False
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindAgain 
         Caption         =   "Find Again"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuWrap 
         Caption         =   "Word Wrap"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "mnuList"
      Visible         =   0   'False
      Begin VB.Menu mnuAnlyzeCgiLink 
         Caption         =   "Anlyze CGI Link"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewWindow 
         Caption         =   "Open In New Window"
      End
      Begin VB.Menu mnuRawReq 
         Caption         =   "Raw Request"
      End
      Begin VB.Menu mnuListPlugin 
         Caption         =   "Plugins"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "Copy Item"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy ALL"
      End
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "mnuFunctions"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCookie 
         Caption         =   "Edit Cookie"
      End
      Begin VB.Menu mnub 
         Caption         =   "HTML Transformations"
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "Embed Script Env"
            Index           =   0
         End
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "hidden -> text"
            Index           =   1
         End
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "select -> text"
            Index           =   2
         End
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "check -> text"
            Index           =   3
         End
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "Radio -> Text"
            Index           =   4
         End
         Begin VB.Menu mnuHTMLTransform 
            Caption         =   "Remove Scripts"
            Index           =   5
         End
      End
      Begin VB.Menu mnuChangeProxy 
         Caption         =   "Change Proxy"
      End
      Begin VB.Menu mnuDirListings 
         Caption         =   "Probe Directories"
      End
      Begin VB.Menu mnuSourceFilter 
         Caption         =   "Set Source Filter"
      End
      Begin VB.Menu mnuViewFrames 
         Caption         =   "Frames Overview"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Generate Report"
      End
   End
   Begin VB.Menu mnuHoldPluginList 
      Caption         =   "Plugins"
      Visible         =   0   'False
      Begin VB.Menu mnuRawRequest 
         Caption         =   "Raw Http Request"
      End
      Begin VB.Menu mnuPlugins 
         Caption         =   "placeholder"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuInternetOpt 
         Caption         =   "Internet Options"
      End
      Begin VB.Menu mnuMouseOverAnlyze 
         Caption         =   "Analyze CGI Links OnMouseOver"
      End
      Begin VB.Menu mnuEditLink 
         Caption         =   "Analyze CGI Links OnNavigate"
      End
      Begin VB.Menu mnuPromptNavigate 
         Caption         =   "Prompt before Navigate"
      End
      Begin VB.Menu mnuAnlyzePost 
         Caption         =   "Analyze POST Data OnSubmit"
      End
      Begin VB.Menu mnuBlockServers 
         Caption         =   "Block Selected Servers"
      End
      Begin VB.Menu mnuNewWinSleuth 
         Caption         =   "New Windows Open w/ Sleuth"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNoPopups 
         Caption         =   "Block Popups"
      End
      Begin VB.Menu mnuLogMode 
         Caption         =   "Log Actions"
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "Bookmarks"
      Visible         =   0   'False
      Begin VB.Menu mnuBookmarkItem 
         Caption         =   "- Add Current Page -"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Web Sleuth: Utility to quickly pick through large web applications
'             and anlyze/ modify forms. Read the read me for usage.
'
'Author: dzzie@yahoo.com

Dim IeDoc As HTMLDocument
Dim filterDelimiter As String
Dim blockServers() As String

Dim Bookmarks() As String
Dim SteppedPath() As String   'holds paths when looking for dir listings
Dim ActionLogFile As String   'path to log file when in log mode
Dim LogTemplate As String     'sets log output format
Dim BookmarkFile As String    'holds all bookmarks
Public EmbedableScriptEnv As String  'path to Included script file


Private Sub Form_Load()
    n = "frmmain"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
    
    With Combo1
     .AddItem "Links", 0:  .AddItem "Images", 1: .AddItem "Frames", 2
     .AddItem "Cookie", 3: .AddItem "Forms", 4:  .AddItem "QueryString Urls", 5
     .AddItem "Scripts", 6: .AddItem "Html Comments", 7
     .AddItem "Meta Tags", 8
    End With
    
    With rtf
        .MatchSize wb: .Top = wb.Top: .Left = wb.Left: .Visible = False
    End With
    
    cmdline = Replace(Command, """", Empty)
    If cmdline = Empty Then
        UseIE_DDE_GETCurrentPageURL txtUrl
        wb.Navigate txtUrl
    Else
        wb.Navigate cmdline
    End If
        
    splitter.ZOrder 0
    EmbedableScriptEnv = App.path & "\Source\EmbedableScriptEnv.html"
    If Not FileExists(EmbedableScriptEnv) Then EmbedableScriptEnv = App.path & "\EmbedableScriptEnv.html"
    
    blockServers() = Split("*doubleclick*,*fusion*,*ad*.com*,*Ads.asp*", ",")
    LogTemplate = vbCrLf & vbCrLf & Date & " %t" & vbCrLf & "Url: %u" & _
                  vbCrLf & "Cookie: %v" & vbCrLf & vbCrLf & String(75, "-")
                  
    BookmarkFile = App.path & "\Bookmarks.txt"
    If FileExists(BookmarkFile) Then
        Dim tmp() As String
        tmp() = Split(ReadFile(BookmarkFile), vbCrLf)
        If Not aryIsEmpty(tmp) Then
            For i = 0 To UBound(tmp)
                If tmp(i) <> Empty Then
                    push Bookmarks(), tmp(i)
                    Load mnuBookmarkItem(i + 1)
                    mnuBookmarkItem(i + 1).caption = ExtractKey(Bookmarks(i))
                End If
            Next
        End If
    End If
    
End Sub

Private Sub cmdEdit_Click()
 On Error GoTo oops
    If Not rtf.Visible Then
       cmdEdit.caption = "View HTML"
       If IeDoc.frames.length > 0 Then
          Set IeDoc = frmFrames.ReturnFrameObject
          If IeDoc Is Nothing Then rtf.text = wb.Document.body.innerHTML: Set IeDoc = wb.Document _
          Else rtf.text = IeDoc.body.innerHTML
       Else
          rtf.text = IeDoc.body.innerHTML
       End If
       rtf.SetSpanColor "<form", ">", vbRed, , True
       rtf.SetColor "</form>", vbRed, , True
       rtf.SetColor "<input", vbBlue, , True
       rtf.SetColor "<script", &HC000C0, , True
       rtf.SetColor "</script>", &HC000C0, , True
       rtf.SetSpanColor "<!--", "-->", &H808000
       rtf.ScrollToTop
       rtf.Visible = True
    Else
       IeDoc.body.innerHTML = CStr(rtf.text) & " "
       Set IeDoc = wb.Document
       cmdEdit.caption = "Edit Source"
       rtf.Visible = False
    End If
Exit Sub
oops: MsgBox Err.Description, vbCritical
End Sub

Private Sub Combo1_Click()
   On Error GoTo oops
   List.Clear
   Select Case Combo1.ListIndex
       Case 0: List.LoadArray GetLinks(IeDoc)
       Case 1: List.LoadArray GetImages(IeDoc)
       Case 2: List.LoadArray GetFrames(IeDoc)
       Case 3: List.LoadArray BreakDownCookie(IeDoc.Cookie)
       Case 4: List.LoadArray GetForms(IeDoc)
       Case 6: List.LoadArray GetScripts(IeDoc)
       Case 7: List.LoadArray GetComments(IeDoc)
       Case 8: List.LoadArray GetMetaTags(IeDoc)
       Case 5:
                Dim tmp()
                Call GetLinks(IeDoc, tmp)
                Call GetImages(IeDoc, tmp)
                Call GetFrames(IeDoc, tmp)
                Call GetScripts(IeDoc, tmp)
                tmp() = List.filterArray(tmp, "*=*")
                If aryIsEmpty(tmp) Then push tmp(), "No QueryString URLs Found in Document"
                List.LoadArray tmp
   End Select
   List.FilterList txtFilter, IIf(lblFilter = "LIKE", True, False)
Exit Sub
oops: MsgBox Err.Description, vbCritical
End Sub

Private Sub mnuBookmarkItem_Click(index As Integer)
   If index = 0 Then
        ans = InputBox("Enter a name to rember this site by", , wb.Document.title)
        push Bookmarks(), ans & "=" & wb.LocationURL
        Load mnuBookmarkItem(mnuBookmarkItem.Count)
        mnuBookmarkItem(mnuBookmarkItem.Count - 1).caption = ans
   Else
        wb.Navigate ExtractValue(Bookmarks(index - 1))
   End If
End Sub

Private Sub mnuReport_Click()
    On Error GoTo warn
    Dim ret(), IncludeSource As Boolean
    IncludeSource = IIf(MsgBox("Do you want to include each docs body.innerHTML ?", vbQuestion + vbYesNo) = vbYes, True, False)
    fpath = App.path & "\Sleuth_Report.txt"
    
    push ret(), vbCrLf & String(75, "-")
    push ret(), Date & String(5, " ") & Time & String(5, " ") & "Saved as: " & fpath & vbCrLf
    push ret(), "If you want to save this file be sure to do a SAVE AS or"
    push ret(), "else it will be automatically overwritten by next report!" & vbCrLf
    
    Call GetPageStats(IeDoc, ret, IncludeSource)
    WriteFile fpath, Join(ret, vbCrLf)
    Shell "notepad """ & fpath & """", vbNormalFocus
    
    Exit Sub
warn: MsgBox Err.Description, vbCritical
End Sub

Private Sub mnuSourceFilter_Click()
    filterDelimiter = InputBox("If the webpage you are viewing uses a huge template that you dont want to scroll through..find a unique string in the document and list it here. Only the portion of the page below that string will be shown while this value is set", , filterDelimiter)
    If filterDelimiter <> Empty Then
        Call ApplyFilter
        If rtf.Visible Then rtf.Visible = False: cmdEdit_Click
    End If
End Sub

Private Sub tmrPauseNavigate_Timer()
    tmrPauseNavigate.Enabled = False
    If UBound(SteppedPath) > 0 Then
        pop SteppedPath
        url = SteppedPath(UBound(SteppedPath))
        wb.Navigate url
    Else
        WipeStrAry SteppedPath
    End If
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    If mnuBlockServers.Checked Then
        For i = 0 To UBound(blockServers)
            If LCase(blockServers(i)) Like LCase(url) Then Cancel = True: Exit Sub
        Next
    End If
    
    If mnuEditLink.Checked Or mnuPromptNavigate.Checked Then
        'this wont modify link browser requests though :(
        url = frmAnalyze.AnlyzeUrl(url, True, Not mnuPromptNavigate.Checked, "Anlyze or Cancel Only, Edit has no effect") & " "
        If url = -1 Then Cancel = True
    End If
    
    If mnuAnlyzePost.Checked And LenB(PostData) > 0 Then
        frmAnalyze.AnlyzeUrl ExtractPostData(PostData), , , "Anlyze Post Data. Note: ReadOnly"
    End If
    
    If mnuLogMode.Checked And LenB(PostData) > 0 Then
        it = vbCrLf & "Data Posted:" & vbCrLf & vbCrLf & Replace(ExtractPostData(PostData), "&", vbCrLf & "&") & vbCrLf & vbCrLf
        If FileExists(ActionLogFile) Then AppendFile ActionLogFile, it _
        Else mnuLogMode.Checked = False
    End If
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, url As Variant)
On Error GoTo oops
  Call ApplyFilter
  Set IeDoc = wb.Document
  Combo1.ListIndex = 0
  Combo1_Click
  txtUrl = wb.Document.location.href
  If mnuLogMode.Checked Then
       it = BatchReplace(LogTemplate, "%t->" & Time & ",%v->" & wb.Document.Cookie & ",%u->" & txtUrl)
        If FileExists(ActionLogFile) Then AppendFile ActionLogFile, it _
        Else mnuLogMode.Checked = False
  End If
  If Not aryIsEmpty(SteppedPath) Then tmrPauseNavigate.Enabled = True
Exit Sub
oops: MsgBox Err.Description, vbCritical
End Sub

Private Sub List_DoubleClick()
  On Error GoTo warn
     If List.Count = 1 And Left(List.SelectedText, 2) = "No" Then Exit Sub
     Select Case Combo1.ListIndex
        Case 0 'links
            t = frmAnalyze.AnlyzeUrl(List.SelectedText, True, False, "Editing Url in Parent Document")
            If t <> -1 Then
                List.UpdateValue t, List.SelectedIndex
                wb.Document.links(List.SelectedIndex).href = t
                Set IeDoc = wb.Document
            End If
        Case 4 'forms data
            If Left(List.value(0), 4) <> "Form" Then
                List.Tag = List.SelectedIndex 'save for if edit latter
                List.LoadArray GetFormsContent(IeDoc, List.SelectedIndex)
            Else
                a = InputBox("Enter new value for this form element")
                IeDoc.Forms(CInt(List.Tag)).elements(CLng(List.SelectedIndex - 3)).value = CStr(a)
                List.LoadArray GetFormsContent(IeDoc, CInt(List.Tag))
            End If
        Case 6 'scripts
            If InStr(List.SelectedText, "Embeded") > 0 Then
                i = List.SelectedIndex
                it = GetScriptContent(IeDoc, i)
                If it <> Empty Then frmAnalyze.ShowScript it
            Else
                If MsgBox("Do you want to do a raw request to view this script?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                mnuRawReq_Click
            End If
        Case 2 'frames
                i = List.SelectedIndex
                db = wb.Document.frames.length
                url = wb.Document.frames(i).location.href
                wb.Navigate url
        Case 3 'cookies
                mnuEditCookie_Click
        Case Else
            frmAnalyze.AnlyzeUrl List.SelectedText, False, False, "Anlyze Url - Note Url will NOT be updated in document"
     End Select
Exit Sub
warn: MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
   If Me.Width > 6700 Then
      wb.Width = Me.Width - 250
      rtf.Width = wb.Width
      Frame1.Width = wb.Width + 50
      splitter.Width = Frame1.Width
      List.Width = wb.Width - 150
      Frame3.Left = wb.Width - Frame3.Width
      txtUrl.Width = Frame3.Left - txtUrl.Left - 150
   End If
   If Me.Height > 5830 Then
      Frame1.Top = Me.Height - Frame1.Height - 400
      wb.Height = Me.Height - Frame1.Height - 350
      splitter.Top = Frame1.Top
      rtf.Height = wb.Height
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    n = "frmmain"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
    End If
    WriteFile BookmarkFile, Join(Bookmarks, vbCrLf)
    Unload frmRawRequest: Unload frmAnalyze: Unload Me
    End
End Sub

Private Sub mnuHTMLTransform_Click(index As Integer)
On Error GoTo oops
   If rtf.Visible Then MsgBox "Sorry only in IE view..or else it reloads the page on change back": Exit Sub
   'menu index corrosponds to the enum index
   If IeDoc.frames.length > 0 Then
        Set IeDoc = frmFrames.ReturnFrameObject
        If IeDoc Is Nothing Then HTMLTransform wb.Document, index _
        Else HTMLTransform IeDoc, index
        Set IeDoc = wb.Document
   Else
        HTMLTransform wb.Document, index
   End If
Exit Sub
oops:  MsgBox Err.Description, vbCritical
End Sub

Private Sub ApplyFilter()
    If filterDelimiter <> Empty Then
        tmp = wb.Document.body.innerHTML
        x = InStr(tmp, filterDelimiter)
        If x > 0 Then wb.Document.body.innerHTML = Mid(tmp, x, Len(tmp)) & " "
    End If
End Sub


'-----------------------------------------------------------------------
'| Misc Events                                                         |
'-----------------------------------------------------------------------
Private Sub Label2_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        Select Case index
            Case 0: PopupMenu mnuFunctions
            Case 1: PopupMenu mnuOptions
            Case 2: PopupMenu mnuHoldPluginList
            Case 3: PopupMenu mnuBookmarks
        End Select
End Sub
Private Sub mnuInternetOpt_Click()
    On Error GoTo oops
    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", vbNormalFocus
oops:
End Sub
Private Sub mnuListPlugin_Click(index As Integer)
     FirePluginEvent index, "frmMain.mnuListPlugin"
End Sub

Private Sub mnuPlugins_Click(index As Integer)
  FirePluginEvent index, "frmMain.mnuPlugins"
End Sub
Private Sub mnuPromptNavigate_Click()
    mnuPromptNavigate.Checked = Not mnuPromptNavigate.Checked
End Sub
Private Sub mnuNewWinSleuth_Click()
 mnuNewWinSleuth.Checked = Not mnuNewWinSleuth.Checked
End Sub
Private Sub List_RightClick()
   PopupMenu mnuList
End Sub
Private Sub rtf_RightClick()
  PopupMenu mnuAdvanced:
End Sub
Private Sub mnuCopyItem_Click()
    Clipboard.Clear: Clipboard.SetText List.SelectedText
End Sub
Private Sub mnuEditLink_Click()
    mnuEditLink.Checked = Not mnuEditLink.Checked
End Sub
Private Sub mnuChangeProxy_Click()
    frmProxy.Show
End Sub
Private Sub mnuFindAgain_Click()
    rtf.findNext
End Sub
Private Sub mnuCopy_Click()
    Clipboard.Clear: Clipboard.SetText List.GetListContents
End Sub
Private Sub mnuNoPopups_Click()
 mnuNoPopups.Checked = Not mnuNoPopups.Checked
End Sub
Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
   If mnuNoPopups.Checked = True Then Cancel = True
End Sub
Private Sub wb_StatusTextChange(ByVal text As String)
    Me.caption = text
    If mnuMouseOverAnlyze.Checked Then frmAnalyze.AnlyzeUrl text
End Sub
Private Sub mnuMouseOverAnlyze_Click()
    mnuMouseOverAnlyze.Checked = Not mnuMouseOverAnlyze.Checked
End Sub
Private Sub mnuAnlyzeCgiLink_Click()
    frmAnalyze.AnlyzeUrl List.SelectedText
End Sub
Private Sub mnuEditCookie_Click()
     frmEditCookie.LoadFormFromUrl wb.Document.location.href
End Sub
Private Sub mnuFind_Click()
    Call rtf.find
End Sub
Sub RenderPage(html)
    wb.Document.body.innerHTML = html & " "
End Sub
Private Sub mnuRawRequest_Click()
    frmRawRequest.PromptForUrlThenInitalize
End Sub
Private Sub lblFilter_Click()
    If lblFilter.caption = "LIKE" Then lblFilter.caption = "NOT" _
    Else lblFilter.caption = "LIKE"
End Sub
Private Sub mnuAnlyzePost_Click()
    mnuAnlyzePost.Checked = Not mnuAnlyzePost.Checked
End Sub
Private Sub txtUrl_DblClick()
    frmAnalyze.AnlyzeUrl txtUrl, False, False, "Examine Url"
End Sub
Private Sub txtFilter_LostFocus()
    Call Combo1_Click 'will auto apply filter after reload list
End Sub
Private Sub txtUrl_GotFocus()
     txtUrl.SelStart = 0
     txtUrl.SelLength = Len(txtUrl.text)
End Sub
Private Sub mnuWrap_Click()
    If mnuWrap.Checked Then rtf.WordWrap = False Else rtf.WordWrap = True
    mnuWrap.Checked = Not mnuWrap.Checked
End Sub
Private Sub txtUrl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If rtf.Visible Then cmdEdit_Click
        wb.Navigate LTrim(Trim(txtUrl))
    End If
End Sub
Private Sub mnuBlockServers_Click()
    mnuBlockServers.Checked = Not mnuBlockServers.Checked
    If mnuBlockServers.Checked Then
        blockServers = Split(InputBox("Enter comma delimited list of wildcard servers to block...Note this is only enabled when this menu item is checked", , Join(blockServers, ",")), ",")
    End If
End Sub
Private Sub mnuDirListings_Click()
    it = Replace(txtUrl, "http://", Empty)
    If it = Empty Or InStr(it, "/") < 1 Then Exit Sub
    SteppedPath() = GetPathsInStep(it)
    wb.Navigate SteppedPath(UBound(SteppedPath))
End Sub
Private Sub mnuNewWindow_Click()
    On Error GoTo out
    StartString = """" & App.path & "\WebSleuth.exe"" """ & frmAnalyze.AnlyzeUrl(List.SelectedText, True, False, "Edit Url to Open In New Sleuth Window") & """"
    Shell StartString, vbNormalFocus
    'Shell "C:\program files\internet explorer\iexplore.exe """ & frmAnalyze.AnlyzeUrl(List.SelectedText & """", True, False, "Edit Url to Open In New Window"), vbNormalFocus
    Exit Sub
out: MsgBox Err.Description, vbExclamation
End Sub
Private Sub mnuViewFrames_Click()
    If wb.Document.frames.length > 0 Then
        frmFrames.ListFrames
    Else
        MsgBox "No frames in current document"
    End If
End Sub
Private Sub cmdNavigate_Click(index As Integer)
    On Error Resume Next
        Select Case index
            Case 0: wb.GoBack
            Case 1: wb.GoForward
            Case 2: wb.Stop
        End Select
End Sub
Private Sub mnuRawReq_Click()
    tmp = List.SelectedText
    If InStr(tmp, "http://") < 1 Then
        tmp = "http://" & wb.Document.location.host & IIf(Left(tmp, 1) = "/", tmp, "/" & tmp)
    End If
    frmRawRequest.PrepareRawRequest tmp, IeDoc.Cookie
End Sub
Private Sub mnuLogMode_Click()
    mnuLogMode.Checked = Not mnuLogMode.Checked
    If mnuLogMode.Checked Then
        ActionLogFile = CmnDlg1.ShowSave(App.path, textFiles, "Write Action Log to")
        If ActionLogFile = Empty Then mnuLogMode.Checked = False: Exit Sub
        WriteFile ActionLogFile, "WebSleuth Surf Report for " & Date & vbCrLf & String(75, "-") & vbCrLf & vbCrLf
    End If
End Sub

Private Sub splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        splitter.Move splitter.Left, (splitter.Top - (splitter.Height \ 2)) + Y
        splitter.BackColor = &H808081
    End If
End Sub

Private Sub splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If splitter.BackColor = &H808081 Then
        splitter.Move splitter.Left, splitter.Top + Y
    End If
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If splitter.BackColor = &H808081 Then
        splitter.BackColor = &H808080
        Frame1.Move Frame1.Left, splitter.Top + Y
        wb.Height = Frame1.Top - wb.Top + 50
        rtf.Height = wb.Height
        Frame1.Height = Me.Height - Frame1.Top - 450
        List.Height = Frame1.Height - List.Top
    End If
End Sub




'----------------------------------------------------------------------
'these are to explose these module functions & objects to plugins
'----------------------------------------------------------------------
Function GetfrmAnalyze() As Object
    Set GetfrmAnalyze = frmAnalyze
End Function

Function getfrmRawReq() As Object
    Set getfrmRawReq = frmRawRequest
End Function

Function getFrmProxy() As Object
    Set getFrmProxy = frmProxy
End Function

Function RemoveHtml(it As String) As String
    RemoveHtml = CStr(parseHtml(it))
End Function

Function URLDecode(it As String) As String
    URLDecode = CStr(UnEscape(it))
End Function

Function URLEncode(it As String, Optional fullEncode As Boolean = True) As String
    URLEncode = CStr(escape(it, fullEncode))
End Function

Function base64Encode(it As String) As String
    base64Encode = b64Encode(it)
End Function

Function base64Decode(it As String) As String
    base64Decode = b64Decode(it)
End Function
