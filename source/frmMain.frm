VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   3840
   End
   Begin TabDlg.SSTab tabBar 
      Height          =   2775
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4895
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   406
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Browser"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "wb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Source"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtf"
      Tab(1).Control(1)=   "frameRTF"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkBlockServers"
      Tab(2).Control(1)=   "chkNoPopups"
      Tab(2).Control(2)=   "chkAnlyzePost"
      Tab(2).Control(3)=   "chkPromptNavigate"
      Tab(2).Control(4)=   "chkEditLink"
      Tab(2).Control(5)=   "chkMouseOverAnlyze"
      Tab(2).Control(6)=   "chkLogMode"
      Tab(2).Control(7)=   "chkUseMouseOvers"
      Tab(2).Control(8)=   "chkDirtySource"
      Tab(2).Control(9)=   "ChkCookieOnTop"
      Tab(2).Control(10)=   "chkSoureFilter"
      Tab(2).Control(11)=   "txtBlockServers"
      Tab(2).Control(12)=   "chkLoadPlugins"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Notes"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ImgNotesSaveAs"
      Tab(3).Control(1)=   "rtfNotes"
      Tab(3).ControlCount=   2
      Begin VB.CheckBox chkLoadPlugins 
         Caption         =   "Load Plugins on Start"
         Height          =   315
         Left            =   -72000
         TabIndex        =   36
         Top             =   1980
         Width           =   1935
      End
      Begin VB.TextBox txtBlockServers 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   2115
         Left            =   -70020
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   35
         Top             =   180
         Width           =   4215
      End
      Begin VB.Frame frameRTF 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   -74880
         TabIndex        =   25
         Top             =   2040
         Width           =   9195
         Begin VB.CheckBox chkWrap 
            Caption         =   "Wrap"
            Height          =   255
            Left            =   6420
            TabIndex        =   34
            ToolTipText     =   "word wrap"
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "UPDATE IE"
            Height          =   315
            Left            =   8100
            TabIndex        =   33
            ToolTipText     =   "Update browser with altered source"
            Top             =   60
            Width           =   1095
         End
         Begin VB.CommandButton cmdHighlight 
            Caption         =   "Colorize"
            Height          =   315
            Left            =   7260
            TabIndex        =   32
            ToolTipText     =   "Defaultsource highlighting scheme"
            Top             =   60
            Width           =   795
         End
         Begin VB.CommandButton cmdReplace 
            Caption         =   "Replace"
            Height          =   315
            Left            =   3120
            TabIndex        =   31
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox txtReplace 
            Height          =   315
            Left            =   2040
            TabIndex        =   30
            ToolTipText     =   "replace the find text with this"
            Top             =   60
            Width           =   1035
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox txtFind 
            Height          =   315
            Left            =   60
            TabIndex        =   28
            ToolTipText     =   "find this text in source"
            Top             =   60
            Width           =   1095
         End
         Begin VB.CommandButton cmdHighLightFindWord 
            Caption         =   "Color Find"
            Height          =   315
            Left            =   5220
            TabIndex        =   27
            ToolTipText     =   "color all instances of word in find box with selected color"
            Top             =   60
            Width           =   1095
         End
         Begin VB.ComboBox cboColor 
            Height          =   315
            ItemData        =   "frmMain.frx":037A
            Left            =   4140
            List            =   "frmMain.frx":038A
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   60
            Width           =   1035
         End
      End
      Begin VB.CheckBox chkSoureFilter 
         Caption         =   "Use Source Filter"
         Height          =   375
         Left            =   -72000
         TabIndex        =   24
         ToolTipText     =   "Removes large templates on fly (to be expanded to more latter)"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox ChkCookieOnTop 
         Caption         =   "Cookie On Top"
         Height          =   315
         Left            =   -72000
         TabIndex        =   23
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CheckBox chkDirtySource 
         Caption         =   "Dirty Source Prompt"
         Height          =   375
         Left            =   -72000
         TabIndex        =   22
         ToolTipText     =   "Remind you to update IE if you edit source"
         Top             =   840
         Width           =   1755
      End
      Begin VB.CheckBox chkUseMouseOvers 
         Caption         =   "Use Mouse Over Navigation Bar"
         Height          =   435
         Left            =   -74820
         TabIndex        =   21
         ToolTipText     =   "Navigation bar eye candy for fancy pants"
         Top             =   1560
         Width           =   2715
      End
      Begin VB.CheckBox chkLogMode 
         Caption         =   "Log Actions"
         Height          =   495
         Left            =   -72000
         TabIndex        =   18
         Top             =   420
         Width           =   1215
      End
      Begin VB.CheckBox chkMouseOverAnlyze 
         Caption         =   "Analyze CGI Links OnMouseOver"
         Height          =   495
         Left            =   -74820
         TabIndex        =   17
         Top             =   60
         Width           =   2835
      End
      Begin VB.CheckBox chkEditLink 
         Caption         =   "Analyze CGI Links OnNavigate"
         Height          =   495
         Left            =   -74820
         TabIndex        =   16
         Top             =   420
         Width           =   2535
      End
      Begin VB.CheckBox chkPromptNavigate 
         Caption         =   "Prompt before Navigate"
         Height          =   495
         Left            =   -74820
         TabIndex        =   15
         Top             =   780
         Width           =   2535
      End
      Begin VB.CheckBox chkAnlyzePost 
         Caption         =   "Analyze POST Data OnSubmit"
         Height          =   495
         Left            =   -74820
         TabIndex        =   14
         Top             =   1140
         Width           =   2655
      End
      Begin VB.CheckBox chkNoPopups 
         Caption         =   "Block Popups"
         Height          =   495
         Left            =   -72000
         TabIndex        =   13
         Top             =   60
         Width           =   1455
      End
      Begin VB.CheckBox chkBlockServers 
         Caption         =   "Block Selected Servers"
         Height          =   495
         Left            =   -74820
         TabIndex        =   12
         ToolTipText     =   "Auto cancel navigation to urls matching reg exp"
         Top             =   1920
         Width           =   2535
      End
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   2400
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   9135
         ExtentX         =   16113
         ExtentY         =   4233
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
      Begin WebSleuth.RTF rtf 
         Height          =   1920
         Left            =   -74880
         TabIndex        =   11
         Top             =   60
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4128
      End
      Begin WebSleuth.RTF rtfNotes 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   19
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4048
      End
      Begin VB.Image ImgNotesSaveAs 
         Height          =   240
         Left            =   -66000
         Picture         =   "frmMain.frx":03A6
         ToolTipText     =   "Save Notes As"
         Top             =   2460
         Width           =   240
      End
   End
   Begin VB.Frame splitter 
      BackColor       =   &H00808080&
      Height          =   60
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   2835
      Width           =   9435
   End
   Begin WebSleuth.CmnDlg CmnDlg1 
      Left            =   9000
      Top             =   3840
      _ExtentX        =   582
      _ExtentY        =   503
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   9465
      Begin WebSleuth.List List 
         Height          =   1470
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "list a range of properties selected with combo box..rt click for tools"
         Top             =   1050
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2593
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
         ItemData        =   "frmMain.frx":0CE7
         Left            =   240
         List            =   "frmMain.frx":0CE9
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "View a specific attribute of the web page"
         Top             =   1080
         Width           =   1710
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
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Addressbar...enter url to surf to and hit return"
         Top             =   240
         Width           =   5505
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   800
         Left            =   6000
         TabIndex        =   4
         Top             =   180
         Width           =   3375
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
            Left            =   1380
            TabIndex        =   5
            Text            =   "*"
            Top             =   480
            Width           =   1920
         End
         Begin WebSleuth.ButtonBar ButtonBar1 
            Height          =   495
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   873
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
            Left            =   720
            TabIndex        =   6
            ToolTipText     =   "click to toggle"
            Top             =   480
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filter"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   60
            TabIndex        =   7
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Image imgSniper 
         Height          =   480
         Left            =   5820
         Picture         =   "frmMain.frx":0CEB
         ToolTipText     =   "Drag & Drop over External IE Window"
         Top             =   540
         Width           =   480
      End
      Begin VB.Image ImgMenuBar 
         Height          =   375
         Index           =   2
         Left            =   2940
         Picture         =   "frmMain.frx":0FF5
         Top             =   600
         Width           =   1410
      End
      Begin VB.Image ImgMenuBar 
         Height          =   375
         Index           =   3
         Left            =   4320
         Picture         =   "frmMain.frx":2000
         Top             =   600
         Width           =   1410
      End
      Begin VB.Image ImgMenuBar 
         Height          =   375
         Index           =   1
         Left            =   1560
         Picture         =   "frmMain.frx":3020
         Top             =   600
         Width           =   1410
      End
      Begin VB.Image ImgMenuBar 
         Height          =   375
         Index           =   0
         Left            =   180
         Picture         =   "frmMain.frx":3FFB
         Top             =   600
         Width           =   1410
      End
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   315
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
      Begin VB.Menu mnuFunction 
         Caption         =   "Advanced"
         Index           =   0
         Begin VB.Menu mnuAdvancedFunction 
            Caption         =   "Internet Options"
            Index           =   0
         End
         Begin VB.Menu mnuAdvancedFunction 
            Caption         =   "Change Proxy"
            Index           =   1
         End
         Begin VB.Menu mnuAdvancedFunction 
            Caption         =   "IE Connect DDE "
            Index           =   2
         End
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "HTML Transformations"
         Index           =   1
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
      Begin VB.Menu mnuFunction 
         Caption         =   "View Full Source"
         Index           =   2
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "Navigate to Frame"
         Index           =   3
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "Frames Overview"
         Index           =   4
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "Generate Report"
         Index           =   5
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
   Begin VB.Menu mnuBookmarks 
      Caption         =   "Bookmarks"
      Visible         =   0   'False
      Begin VB.Menu mnuBookmarkItem 
         Caption         =   "- Add Current Page -"
         Index           =   0
      End
   End
   Begin VB.Menu mnuExtactFx 
      Caption         =   "mnuExtactFx"
      Visible         =   0   'False
      Begin VB.Menu mnuExtact 
         Caption         =   "Links"
         Index           =   0
      End
      Begin VB.Menu mnuExtact 
         Caption         =   "Forms"
         Index           =   1
      End
      Begin VB.Menu mnuExtact 
         Caption         =   "Cookie"
         Index           =   2
      End
      Begin VB.Menu mnuExtact 
         Caption         =   "Frames"
         Index           =   3
      End
      Begin VB.Menu mnuExtact 
         Caption         =   "QueryStrings"
         Index           =   4
      End
      Begin VB.Menu mnuExtact 
         Caption         =   "Extra"
         Index           =   5
         Begin VB.Menu mnuExtraFx 
            Caption         =   "Images"
            Index           =   0
         End
         Begin VB.Menu mnuExtraFx 
            Caption         =   "Scripts"
            Index           =   1
         End
         Begin VB.Menu mnuExtraFx 
            Caption         =   "Comments"
            Index           =   2
         End
         Begin VB.Menu mnuExtraFx 
            Caption         =   "Meta Tags"
            Index           =   3
         End
         Begin VB.Menu mnuExtraFx 
            Caption         =   "Emails"
            Index           =   4
         End
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

Dim ieDoc As HTMLDocument     'this is most often set to be the current document, but is also passed off to frames etc at times
Dim filterDelimiter As String 'to parse out unintresting templates kinda obsolete now with partial editing capability
Dim LogTemplate As String     'sets log output format
Dim partialEdit As Boolean    'if we are editing only part of page <--cool thanks to thePull for implementation ideas

Private Sub Form_Load()
    n = "frmmain"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
    
    chkUseMouseOvers.value = GetSetting(App.title, "Options", "Mousie", 1)
    chkNoPopups.value = GetSetting(App.title, "Options", "Popups", 0)
    chkDirtySource.value = GetSetting(App.title, "Options", "DirtySource", 1)
    chkLoadPlugins.value = GetSetting(App.title, "Options", "LoadPlugins", 1)
    
    With Combo1
     .AddItem "Links", 0:  .AddItem "Images", 1: .AddItem "Frames", 2
     .AddItem "Cookie", 3: .AddItem "Forms", 4:  .AddItem "QueryString Urls", 5
     .AddItem "Scripts", 6: .AddItem "Html Comments", 7
     .AddItem "Meta Tags", 8
    End With
    
    Call SetGlobalFilePaths
    cboColor.ListIndex = 0
    splitter.ZOrder 0
    tabBar.Tab = 0
    
    cmdline = Replace(Command, """", Empty)
    If cmdline = Empty Then
        UseIE_DDE_GETCurrentPageURL txtUrl
        wb.Navigate txtUrl
    Else
        wb.Navigate cmdline
    End If
        
    LogTemplate = vbCrLf & vbCrLf & Date & " %t" & vbCrLf & "Url: %u" & _
                  vbCrLf & "Cookie: %v" & vbCrLf & vbCrLf & String(75, "-")
    

End Sub

Sub SetRtfText(d As HTMLDocument)
    Dim sel As String
    sel = GetSelectedHtml(d)
    If Len(sel) > 0 Then
        partialEdit = True
        rtf.text = sel
    Else
        rtf.text = d.body.innerHTML
    End If
End Sub

Private Sub ImgNotesSaveAs_Click()
    On Error Resume Next
    ret = CmnDlg1.ShowSave(App.path, textFiles, "Save Notes As:", True)
    If ret = Empty Then Exit Sub
    WriteFile ret, rtfNotes.text
End Sub


Private Sub tabBar_Click(PreviousTab As Integer)
On Error GoTo oops

    If PreviousTab = 3 Then 'leaving Notes tab so save changes
        WriteFile NotesFile, rtfNotes.text
    End If
        
    If PreviousTab = 1 And rtf.IsDirty And CBool(chkDirtySource.value) Then
        If MsgBox("Html source has been altered do you wish to change rendered page?", vbInformation + vbYesNo) = vbYes Then
            cmdUpdate_Click
        End If
    End If
        
    If tabBar.Tab = 1 And PreviousTab = 0 Then
       'they are coming from browser to edit source
       If frmFrames.AnyAccessibleFrames Then
          Set ieDoc = frmFrames.ReturnFrameObject
          If ieDoc Is Nothing Then
                SetRtfText wb.Document
                Set ieDoc = wb.Document
          Else
                SetRtfText ieDoc
          End If
       Else
          SetRtfText ieDoc
       End If
    End If
    
Exit Sub
oops: MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdUpdate_Click() 'update browser source from edit
 On Error GoTo oops
      If partialEdit Then
           objSelection.pasteHTML rtf.text
           objSelDocument.selection.Clear
           Set objSelection = Nothing
           Set objSelDocument = Nothing
           partialEdit = False
       Else
           ieDoc.body.innerHTML = CStr(rtf.text) & " "
       End If
       Set ieDoc = wb.Document
       rtf.IsDirty = False
       tabBar.Tab = 0
Exit Sub
oops: MsgBox Err.Description, vbCritical
End Sub

Private Sub mnuExtact_Click(index As Integer)
    'this is just bodgered together for now to see if i dig no combobox
    On Error GoTo oops
    Select Case index
        Case 0 'Extract Links
                Combo1.ListIndex = 0
        Case 1 'Forms
                Combo1.ListIndex = 4
        Case 2 'Cookies
                Combo1.ListIndex = 3
        Case 3 'Frames
                Combo1.ListIndex = 2
        Case 4 'QueryStrings
                Combo1.ListIndex = 5
    End Select
oops:
End Sub

Private Sub mnuExtraFx_Click(index As Integer)
    Select Case index
        Case 0 'images
                Combo1.ListIndex = 1
        Case 1 'scripts
                Combo1.ListIndex = 6
        Case 2 'comments
                Combo1.ListIndex = 7
        Case 3 'meta
                Combo1.ListIndex = 8
        Case 4 'emails
                MsgBox "Coming soon to a sleuth near you", vbInformation
    End Select
End Sub
Private Sub Combo1_Click()
   On Error GoTo oops
   List.Clear
   mnuRawReq.Visible = IIf(Combo1.ListIndex < 3, True, False)
   
   If Combo1.ListIndex = 4 Then
        mnuRawReq.Visible = True
        mnuRawReq.Tag = "can initalize post from form <--i am fancy dancy!"
   Else
        mnuRawReq.Tag = Empty
   End If
   
   Select Case Combo1.ListIndex
       Case 0: List.LoadArray GetLinks(ieDoc)
       Case 1: List.LoadArray GetImages(ieDoc)
       Case 2: List.LoadArray GetFrames(ieDoc)
       Case 3: List.LoadArray BreakDownCookie(ieDoc.cookie)
       Case 4: List.LoadArray GetForms(ieDoc)
       Case 6: List.LoadArray GetScripts(ieDoc)
       Case 7: List.LoadArray GetComments(ieDoc)
       Case 8: List.LoadArray GetMetaTags(ieDoc)
       Case 5:
                Dim tmp()
                Call GetLinks(ieDoc, tmp)
                Call GetImages(ieDoc, tmp)
                Call GetFrames(ieDoc, tmp)
                Call GetScripts(ieDoc, tmp)
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
        If ans = Empty Then MsgBox "Ughh no blanks sorry": Exit Sub
        push Bookmarks(), ans & "=" & wb.LocationURL
        Load mnuBookmarkItem(mnuBookmarkItem.Count)
        mnuBookmarkItem(mnuBookmarkItem.Count - 1).caption = ans
   Else
        wb.Navigate ExtractValue(Bookmarks(index - 1))
   End If
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    If CBool(chkBlockServers.value) Then
        For i = 0 To UBound(blockServers)
            If LCase(blockServers(i)) Like LCase(url) Then Cancel = True: Exit Sub
        Next
    End If
    
    If CBool(chkEditLink.value) Or CBool(chkPromptNavigate.value) Then
        'this wont modify link browser requests though :(
        url = frmAnalyze.AnlyzeUrl(url, True, Not CBool(chkPromptNavigate.value), "Anlyze or Cancel Only, Edit has no effect") & " "
        If url = -1 Then Cancel = True
    End If
    
    If CBool(chkAnlyzePost.value) And LenB(PostData) > 0 Then
        frmAnalyze.AnlyzeUrl ExtractPostData(PostData), , , "Anlyze Post Data. Note: ReadOnly"
    End If
    
    If CBool(chkLogMode.value) And LenB(PostData) > 0 Then
        it = vbCrLf & "Data Posted:" & vbCrLf & vbCrLf & Replace(ExtractPostData(PostData), "&", vbCrLf & "&") & vbCrLf & vbCrLf
        If FileExists(ActionLogFile) Then AppendFile ActionLogFile, it _
        Else chkLogMode.value = 0
    End If
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, url As Variant)
On Error GoTo oops
  Call ApplyFilter
  Set ieDoc = wb.Document
  Combo1.ListIndex = 0
  Combo1_Click
  txtUrl = wb.Document.location.href
  If CBool(chkLogMode.value) Then
       it = BatchReplace(LogTemplate, "%t->" & Time & ",%v->" & wb.Document.cookie & ",%u->" & txtUrl)
        If FileExists(ActionLogFile) Then AppendFile ActionLogFile, it _
        Else chkLogMode.value = 0
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
                Set ieDoc = wb.Document
            End If
        Case 4 'forms data
            If Left(List.value(0), 4) <> "Form" Then
                mnuRawReq.Visible = False
                mnuRawReq.Tag = Empty
                List.Tag = List.SelectedIndex 'save for if edit latter
                List.LoadArray GetFormsContent(ieDoc, List.SelectedIndex)
            Else
                oldVal = ieDoc.Forms(CInt(List.Tag)).elements(CLng(List.SelectedIndex - 3)).value
                a = InputBox("Enter new value for this form element", , oldVal)
                ieDoc.Forms(CInt(List.Tag)).elements(CLng(List.SelectedIndex - 3)).value = CStr(a)
                List.LoadArray GetFormsContent(ieDoc, CInt(List.Tag))
            End If
        Case 6 'scripts
            If InStr(List.SelectedText, "SRC = ") < 1 Then
                frmAnalyze.ShowScript List.SelectedText
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
                mnuFunction_Click (0)
        Case Else
            frmAnalyze.AnlyzeUrl List.SelectedText, False, False, "Anlyze Url - Note Url will NOT be updated in document"
     End Select
Exit Sub
warn: MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
   If Me.Width > 6700 Then
      tabBar.Width = Me.Width - 150
      wb.Width = tabBar.Width - 150
      rtf.Width = wb.Width
      rtfNotes.Width = wb.Width
      ImgNotesSaveAs.Left = rtfNotes.Left + rtfNotes.Width - ImgNotesSaveAs.Width - 70
      Frame1.Width = tabBar.Width + 50
      splitter.Width = Frame1.Width
      List.Width = wb.Width - 150
      Frame3.Left = wb.Width - Frame3.Width
      txtUrl.Width = Frame3.Left - txtUrl.Left - 150
   End If
   If Me.Height > 5830 Then
      Frame1.Top = Me.Height - Frame1.Height - 400
      tabBar.Height = Me.Height - Frame1.Height - 500
      wb.Height = tabBar.Height - 425
      rtfNotes.Height = wb.Height
      ImgNotesSaveAs.Top = rtfNotes.Top + rtfNotes.Height + 17
      splitter.Top = Frame1.Top
      rtf.Height = wb.Height - frameRTF.Height - 200
      frameRTF.Top = rtf.Top + rtf.Height + 120
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
    SaveSetting App.title, "Options", "Mousie", chkUseMouseOvers.value
    SaveSetting App.title, "Options", "Popups", chkNoPopups.value
    SaveSetting App.title, "Options", "DirtySource", chkDirtySource.value
    SaveSetting App.title, "Options", "LoadPlugins", chkLoadPlugins.value
    WriteFile BookmarkFile, Join(Bookmarks, vbCrLf)
    WriteFile BlockServerFile, Join(blockServers, vbCrLf)
    Set ieDoc = Nothing
    For Each f In VB.Forms: Unload f: Next
    End
End Sub

Private Sub mnuHTMLTransform_Click(index As Integer)
On Error GoTo oops
   If tabBar.Tab > 0 Then MsgBox "Sorry only in IE view..or else it reloads the page on change back": Exit Sub
   'menu index corrosponds to the enum index
   If frmFrames.AnyAccessibleFrames Then
        Set ieDoc = frmFrames.ReturnFrameObject
        If ieDoc Is Nothing Then HTMLTransform wb.Document, index _
        Else HTMLTransform ieDoc, index
        Set ieDoc = wb.Document
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

Private Sub mnuAdvancedFunction_Click(index As Integer)
On Error GoTo oops
    Select Case index
        Case 0 'Internet Options
                Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", vbNormalFocus
        Case 1 'Change Proxy
                frmProxy.Show
        Case 2 'Ie Connect
                before = txtUrl
                UseIE_DDE_GETCurrentPageURL txtUrl
                If txtUrl <> before Then wb.Navigate txtUrl
    End Select
Exit Sub
oops:  MsgBox Err.Description, vbCritical
End Sub

Private Sub mnuFunction_Click(index As Integer)
 On Error GoTo oops
    Select Case index
        Case 2 'View Full Source
                wb.Navigate "view-source:" & txtUrl
        Case 3: 'navigate to frame
                If frmFrames.AnyAccessibleFrames Then
                    Dim tmpDoc As HTMLDocument
                    Set tmpDoc = frmFrames.ReturnFrameObject
                    If tmpDoc Is Nothing Then Exit Sub
                    t = tmpDoc.location.href
                    wb.Navigate2 t
                End If
        Case 4 'Frames Overview
                If wb.Document.frames.length > 0 Then frmFrames.ListFrames _
                Else MsgBox "No frames in current document"
        Case 5 'Gen Report
                Call GenReport
    End Select
Exit Sub
oops:  MsgBox Err.Description, vbCritical
End Sub
'-----------------------------------------------------------------------
'| Misc Events                                                         |
'-----------------------------------------------------------------------
Private Sub ImgMenuBar_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        Select Case index
            Case 0: PopupMenu mnuExtactFx
            Case 1: PopupMenu mnuFunctions
            Case 2: PopupMenu mnuHoldPluginList
            Case 3: PopupMenu mnuBookmarks
        End Select
End Sub

Private Sub chkLoadPlugins_Click()
    If chkLoadPlugins.Tag = "Block Repete from autoset" Then
        chkLoadPlugins.Tag = Empty
        Exit Sub
    End If
    
    If tabBar.Tab = 2 And aryIsEmpty(RegisteredPlugins) Then
        If MsgBox("Would you like to load plugins now?", vbInformation + vbYesNo) = vbYes Then
            LoadPlugins
        Else
            chkLoadPlugins.Tag = "Block Repete from autoset"
            chkLoadPlugins.value = 0
        End If
    End If
End Sub

Private Sub chkSoureFilter_Click()
    filterDelimiter = InputBox("If the webpage you are viewing uses a huge template that you dont want to scroll through..find a unique string in the document and list it here. Only the portion of the page below that string will be shown while this value is set", , filterDelimiter)
    If filterDelimiter <> Empty Then
        chkSoureFilter.value = 1
        Call ApplyFilter
        If tabBar.Tab > 0 Then tabBar.Tab = 0
    Else
        chkSoureFilter.value = 0
    End If
End Sub

Private Sub chkUseMouseOvers_Click()
    ButtonBar1.RunTimeStyle = chkUseMouseOvers.value
End Sub

Private Sub cmdHighLightFindWord_Click()
    LockWindowUpdate rtf.hWnd
    c = Array(vbRed, vbBlue, &HC000C0, &H808000)
    rtf.SetColor txtFind, CLng(c(cboColor.ListIndex)), , True
    rtf.ScrollToTop
    LockWindowUpdate 0&
End Sub

Private Sub txtBlockServers_LostFocus()
    blockServers() = Split(txtBlockServers, vbCrLf)
End Sub

Private Sub mnuListPlugin_Click(index As Integer)
     FirePluginEvent index, "frmMain.mnuListPlugin"
End Sub

Private Sub mnuPlugins_Click(index As Integer)
  FirePluginEvent index, "frmMain.mnuPlugins"
End Sub
Private Sub List_RightClick()
   PopupMenu mnuList
End Sub

Private Sub mnuCopyItem_Click()
    Clipboard.Clear: Clipboard.SetText List.SelectedText
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear: Clipboard.SetText List.GetListContents
End Sub

Private Sub wb_NavigateComplete2(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    If CBool(ChkCookieOnTop.value) Then
        frmMmmmCookies.ShowCookie wb.Document.cookie
    End If
End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
   If CBool(chkNoPopups.value) = True Then Cancel = True
End Sub
Private Sub wb_StatusTextChange(ByVal text As String)
    Me.caption = text
    If CBool(chkMouseOverAnlyze.value) Then frmAnalyze.AnlyzeUrl text
End Sub
Private Sub chkWrap_Click()
    rtf.WordWrap = CBool(chkWrap.value)
End Sub

Private Sub cmdFind_Click()
    If txtFind <> rtf.FindString Then
        rtf.FindString = txtFind
        rtf.find
    Else
        rtf.findNext
    End If
End Sub

Private Sub cmdReplace_Click()
    rtf.ReplaceText txtFind, txtReplace
End Sub

Private Sub mnuAnlyzeCgiLink_Click()
    frmAnalyze.AnlyzeUrl List.SelectedText
End Sub

Private Sub mnuRawRequest_Click()
    frmRawRequest.PromptForUrlThenInitalize
End Sub
Private Sub lblFilter_Click()
    If lblFilter.caption = "LIKE" Then lblFilter.caption = "NOT" _
    Else lblFilter.caption = "LIKE"
End Sub

Private Sub txtUrl_DblClick()
    frmAnalyze.AnlyzeUrl txtUrl, False, False, "Examine Url"
End Sub
Private Sub txtFilter_LostFocus()
    Call Combo1_Click 'will auto apply filter after reload list
End Sub
Private Sub ChkCookieOnTop_Click()
    frmMmmmCookies.Visible = CBool(ChkCookieOnTop.value)
End Sub

Private Sub txtUrl_GotFocus()
     txtUrl.SelStart = 0
     txtUrl.SelLength = Len(txtUrl.text)
End Sub
Private Sub cmdHighlight_Click()
    LockWindowUpdate rtf.hWnd
    rtf.highlightHtml
    LockWindowUpdate 0&
End Sub

Private Sub txtUrl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If tabBar.Tab <> 0 Then tabBar.Tab = 0
        wb.Navigate LTrim(Trim(txtUrl))
    End If
End Sub

Private Sub chkBlockServers_Click()
    txtBlockServers.Enabled = CBool(chkBlockServers.value)
End Sub

Private Sub mnuNewWindow_Click()
    On Error GoTo out
    StartString = """" & App.path & "\WebSleuth.exe"" """ & frmAnalyze.AnlyzeUrl(List.SelectedText, True, False, "Edit Url to Open In New Sleuth Window") & """"
    Shell StartString, vbNormalFocus
    'Shell "C:\program files\internet explorer\iexplore.exe """ & frmAnalyze.AnlyzeUrl(List.SelectedText & """", True, False, "Edit Url to Open In New Window"), vbNormalFocus
    Exit Sub
out: MsgBox Err.Description, vbExclamation
End Sub

Private Sub ButtonBar1_Click(index As Integer)
    On Error Resume Next
        tabBar.Tab = 0
        Select Case index
            Case 1: wb.Navigate LTrim(Trim(txtUrl))
            Case 0: wb.GoBack
            Case 4: wb.GoForward
            Case 2: wb.Stop
            Case 3: wb.Refresh2 9
        End Select
End Sub

Private Sub mnuRawReq_Click()
 On Error GoTo shit
    'remove SRC = in case coming from script src entry
    tmp = Replace(List.SelectedText, "SRC = ", Empty)
    If mnuRawReq.Tag <> Empty Then
        'initalize raw request from form data and set req to be post
        formAsQueryString = TurnFormIntoQueryString(wb.Document, List.SelectedIndex)
        usePOST = IIf(wb.Document.Forms(List.SelectedIndex).method = "GET", False, True)
        frmRawRequest.PrepareRawRequest formAsQueryString, ieDoc.cookie, usePOST
    Else
        If InStr(tmp, "http://") < 1 Then
            tmp = "http://" & wb.Document.location.host & IIf(Left(tmp, 1) = "/", tmp, "/" & tmp)
        End If
        frmRawRequest.PrepareRawRequest tmp, ieDoc.cookie
    End If
 Exit Sub
shit:  MsgBox "I tried but this is hard :(" & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub chkLogMode_Click()
    If CBool(chkLogMode.value) Then
        ActionLogFile = CmnDlg1.ShowSave(App.path, textFiles, "Write Action Log to")
        If ActionLogFile = Empty Then chkLogMode.value = 0: Exit Sub
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
        tabBar.Height = Frame1.Top - tabBar.Top - 125
        wb.Height = tabBar.Height - 425
        rtf.Height = wb.Height - frameRTF.Height - 200
        frameRTF.Top = rtf.Top + rtf.Height + 120
        Frame1.Height = Me.Height - Frame1.Top - 450
        List.Height = Frame1.Height - List.Top
        Form_Resize 'this could probably replace most of these here duh i am dumb
    End If
End Sub




'----------------------------------------------------------------------
'these are to expose these module functions & objects to plugins
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

Sub TransferObject(d As HTMLDocument)
    Set ieDoc = d
    txtUrl = ieDoc.location
    wb.Document.body.innerHTML = ieDoc.body.innerHTML
    Combo1.ListIndex = 1 'necessary to get list to update
    Combo1.ListIndex = 0
End Sub

Private Sub imgSniper_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Screen.MousePointer = 99 'custom
    Screen.MouseIcon = LoadResPicture("sniper.ico", vbResIcon)
    Timer1.Enabled = True
End Sub

Private Sub imgSniper_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Timer1.Enabled = False
    Screen.MousePointer = vbDefault
    DoEvents
    If modSniper.IsIEServerWindow(CLng(Me.caption)) Then
        TransferObject modSniper.IEDOMFromhWnd(CLng(Me.caption))
    Else
        Me.caption = "Not Valid IE Window"
    End If
End Sub

Private Sub Timer1_Timer()
    Dim p As POINTAPI
    GetCursorPos p
    Me.caption = WindowFromPoint(p.x, p.Y)
End Sub
