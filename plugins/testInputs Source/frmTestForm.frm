VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmTestForm 
   Caption         =   "Test Form Inputs Plugin"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   330
      Left            =   2310
      TabIndex        =   15
      Top             =   2415
      Width           =   960
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   330
      Left            =   1365
      TabIndex        =   14
      Top             =   2415
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Manage Test List"
      Height          =   330
      Left            =   3255
      TabIndex        =   13
      Top             =   2415
      Width           =   1485
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6510
      TabIndex        =   12
      Text            =   "1"
      Top             =   2415
      Width           =   330
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6405
      Top             =   1890
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2535
      Left            =   105
      TabIndex        =   10
      Top             =   2835
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   4471
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
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5460
      TabIndex        =   8
      Text            =   "3"
      Top             =   2415
      Width           =   330
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Left            =   6405
      Top             =   1470
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Index           =   1
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   420
      Width           =   4110
   End
   Begin VB.ListBox List1 
      Height          =   1410
      Left            =   105
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   945
      Width           =   6735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTestForm.frx":0000
      Left            =   5250
      List            =   "frmTestForm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Index           =   0
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   105
      Width           =   5790
   End
   Begin VB.CommandButton cmdBeginTest 
      Caption         =   "Begin Test"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   2415
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Delay"
      Height          =   195
      Index           =   1
      Left            =   5985
      TabIndex        =   11
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Timeout"
      Height          =   195
      Index           =   0
      Left            =   4830
      TabIndex        =   9
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Parent Href: "
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "All Checked Form Elements will be tested"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   735
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Form Action:"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   960
   End
End
Attribute VB_Name = "frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private formIndex As Integer
Private parentHref As String
Private timeOutOveride As Boolean
Private WaitingToReturn As Boolean
Private pause As Boolean
Private stopIt As Boolean
Private report()

Private Sub BeginTest()
    tmrTimeout.Interval = Text2 * 1000
    tmrTimeout.Enabled = True
    If pause Then cmdPause_Click
    stopIt = False
    
    ReDim report(0)
    report(0) = "Test for: " & parentHref & " Form Index: " & formIndex & vbCrLf & vbCrLf
    
    For i = 0 To wb.Document.Forms(formIndex).Length - 1
        t = UCase(wb.Document.Forms(formIndex).elements(i).Type)
        n = wb.Document.Forms(formIndex).elements(i).Name
        If Not List1.Selected(i) Or AnyOfTheseInstr(t, "BUTTON,RESET,SUBMIT") Then GoTo nextOne
        For j = 0 To UBound(testString)
            If stopIt Then GoTo gameOver
            Call SetDefaultValues
            hit = ExtractCommands(testString(j))
            Select Case hit(0)
                Case "QUIT": GoTo gameOver
                Case "GOTO": j = hit(1): GoTo innerSkip
                Case "NEXT": GoTo nextOne
                Case Empty:  GoTo innerSkip
            End Select
            hit(0) = autoExpand(hit(0))    'expand buffer check notation
            If InStr(hit(0), "[") > 0 Then 'testing a range of values
                Call ScanRange(hit, n)
            Else
                ret = AlterFormSubmitGOBack(i, hit(0))
                Call preformCriteriaTest(ret, hit(0), hit(1), hit(2), n)
            End If
innerSkip:
        Next
nextOne:
     Next
gameOver:
     frmAnalyze.ShowScript Join(report, vbCrLf)
End Sub

Function ExtractCommands(it)
    'INJECT cmd returns insertString,lookfor,logmsg
    'IF cmd returns GOTO,lineNumber or NEXT (next element in form)
    'GOTO x valid as line as label x is in script
    'GOTO NEXT skips to next form element to test
    'QUIT also valid so you can terminate jumps
    'anything else returns empty which will be ignored
    'comments are allowed..lines start with a #
    Dim cmd()
    Dim ret(3)
    cmd() = cmdline.GetArgs(it)
    If Left(cmd(0), 1) = "#" Then GoTo returnNow

    Select Case UCase(cmd(0))
        Case "INJECT", "INSERT"
            ret(0) = cmd(1)
            ret(1) = cmd(3)
            ret(2) = cmd(5)
        Case "EXIT", "QUIT", "END"
            ret(0) = "QUIT"
        Case "GOTO"
             If UCase(cmd(5)) = "NEXT" Then
                ret(0) = "NEXT"
             Else
                tmp = LineNumberFromLabel(cmd(5))
                ret(0) = IIf(tmp = Empty, "QUIT", "GOTO")
                ret(1) = tmp
             End If
        Case "IF"
            Dim find As String, found As Boolean
            If UCase(cmd(1)) = "NOT" Then find = cmd(2) Else find = cmd(1)
            
            For k = 0 To UBound(report)
               If InStr(1, report(k), find, 1) > 0 Then found = True: Exit For
            Next
            
            If UCase(cmd(1)) = "NOT" Then found = Not found
            If Not found Then GoTo returnNow
            
            Select Case UCase(cmd(4))
                Case "GOTO"
                    If UCase(cmd(5)) = "NEXT" Then
                       ret(0) = "NEXT"
                    Else
                       tmp = LineNumberFromLabel(cmd(5))
                       ret(0) = IIf(tmp = Empty, "QUIT", "GOTO")
                       ret(1) = tmp
                    End If
                Case "QUIT", "END", "EXIT"
                    ret(0) = "QUIT"
            End Select
    End Select
    ExtractCommands = ret()
Exit Function
returnNow: ExtractCommands = ret()
End Function

Function LineNumberFromLabel(labelName)
    Dim r
    For k = 0 To UBound(testString)
        If Trim(LTrim(testString(k))) = labelName Then
            r = k     'set equal to teststring index
            Exit For
        End If
    Next
    If r = Empty Then MsgBox "Could not Find Label '" & labelName & "'" & vbCrLf & vbCrLf & "Exiting Script"
    LineNumberFromLabel = r
End Function

Function autoExpand(it)
    If Left(it, 3) Like "*\x" Then
        c = Left(it, 1)
        X = Mid(it, 4, Len(it))
        autoExpand = String(X, c)
    Else
        autoExpand = it
    End If
End Function

Sub SetDefaultValues()
  On Error GoTo oops
    For i = 0 To List1.ListCount - 1
        t = List1.List(i)
        If Left(t, 7) = "DEFAULT" Then
            v = Mid(t, InStr(t, "=") + 1, Len(t))
            wb.Document.Forms(formIndex).elements(i).value = v
        End If
    Next
    Exit Sub
oops: MsgBox "Err in SetDefaultValue: " & Err.Description
End Sub

Function ScanRange(hit, ElementName)
        lb = InStr(hit(0), "[")
        rb = InStr(hit(0), "]")
        prefix = Mid(hit(0), 1, lb - 1)
        suffix = Mid(hit(0), rb + 1, Len(hit(0)))
        range = Split(Mid(hit(0), lb + 1, rb - 1 - lb), "-")
        
        Dim usebuffer As Boolean, buffer As String
        If Left(range(0), 1) = 0 And Len(range(0)) > 1 Then usebuffer = True
        If UBound(range) = 3 Then range(2) = -range(2)
        
        For k = range(0) To range(1) Step range(2)
            If usebuffer Then
                lbuf = Len(range(1)) - Len(k)
                If lbuf > 0 Then buffer = String(lbuf, "0") Else buffer = Empty
            End If
            rtest = prefix & buffer & k & suffix
            ret = AlterFormSubmitGOBack(i, rtest)
            Call preformCriteriaTest(ret, rtest, hit(1), hit(2), ElementName)
            If stopIt Then Exit For
        Next
End Function

Function AlterFormSubmitGOBack(ElementIndex, NewValue)
    wb.Document.Forms(formIndex).elements(ElementIndex).value = NewValue
    wb.Document.Forms(formIndex).submit
    ret = WaitforDocumentComplete
    wb.Navigate2 parentHref
    WaitforDocumentComplete
    AlterFormSubmitGOBack = ret
End Function

Sub preformCriteriaTest(pageContents, testStr, findWhat, reportValue, ElementName)
    save = False
    If UCase(findWhat) = "IT" Then 'unaltered string
        If InStr(1, pageContents, testStr, vbTextCompare) > 0 Then save = True
    ElseIf UCase(findWhat) = "SQLERROR" Then
        If wasBadSqlError(pageContents) Then save = True
    Else 'custom criteria
        If InStr(1, pageContents, findWhat, vbTextCompare) > 0 Then save = True
    End If
    If save Then push report(), "Input: '" & ElementName & "' -->  " & reportValue
End Sub



'-----------------------------------------------------------------------
'Form Setup and Initalization
'-----------------------------------------------------------------------
Sub InitalizeFromForm(url, Index)
   On Error GoTo out
    parentHref = url
    formIndex = Index
    wb.Navigate2 url
    WaitforDocumentComplete
    List1.Clear
    With wb.Document.Forms(formIndex)
        Text1(0) = wb.Document.location.href
        Text1(1) = .Action
        Combo1.ListIndex = IIf(UCase(.method) = "GET", 0, 1)
        For j = 0 To .elements.Length - 1
            t = UCase(.elements(j).Type)
            List1.AddItem t & " - " & .elements(j).Name & "=" & .elements(j).value
            If t <> "BUTTON" And t <> "RESET" And t <> "SUBMIT" Then List1.Selected(j) = True
        Next
    End With
    Me.Visible = True
    Exit Sub
out: MsgBox "Err Loading form..couldnt find form index:" & Index & vbCrLf & vbCrLf & "In Url:" & url, , Err.Description
     Me.Visible = False
End Sub

Sub InitalizeFromUrl(url)
    Text1(0) = url
    url = Replace(url, "http://", Empty)
    s = InStr(url, "/")
    q = InStr(url, "?")
    If s > 0 And q > 0 Then
        server = Mid(url, 1, s - 1)
        Page = Mid(url, s, q - s)
        qs = Mid(url, q + 1, Len(url))
        frm = Split(qs, "&")
        Dim tmp()
        push tmp(), "<html><body><form method=get action=""http://" & server & Page & """><b><font size=+2>"
        For i = 0 To UBound(frm)
            n = Split(frm(i), "=")
            push tmp(), n(0) & "<input type=text name='" & n(0) & "' value='" & n(1) & "'><br>"
        Next
        push tmp(), "<input type=submit value=submit></form>"
        WriteFile App.path & "\form.html", Join(tmp, vbCrLf)
        'Shell "notepad """ & App.path & "\form.html""", vbNormalFocus
        Me.Visible = True
        InitalizeFromForm App.path & "\form.html", 0
    Else
        MsgBox "This is for scanning a script from a CGI Url..no query string detected!", vbExclamation
        Me.Visible = False
    End If
End Sub



'-----------------------------------------------------------------------
'Delay & Wait Subs
'-----------------------------------------------------------------------
Function WaitforDocumentComplete() As String
    WaitingToReturn = True
    While WaitingToReturn Or pause
        DoEvents: DoEvents
        If timeOutOveride Then WaitingToReturn = False: 'MsgBox "Timeout!"
    Wend
    tmrTimeout.Enabled = False
    DelayFor txtDelay
    tmp = Replace(wb.Document.body.innerHTML, vbCrLf, Empty)
    WaitforDocumentComplete = Replace(tmp, vbLf, Empty)
End Function

Function DelayFor(xSecs)
    tmrDelay.Interval = xSecs * 1000
    tmrDelay.Enabled = True
    WaitingToReturn = True
    While WaitingToReturn
        DoEvents: DoEvents
    Wend
    tmrDelay.Enabled = False
End Function

Function WaitForResume()
    WaitingToReturn = True
    While WaitingToReturn
        DoEvents
    Wend
End Function





'-----------------------------------------------------------------------
'Form Events
'-----------------------------------------------------------------------
Private Sub cmdBeginTest_Click()
        Call BeginTest
End Sub

Private Sub cmdPause_Click()
    If pause Then
        cmdPause.Caption = "Pause"
        pause = False
    Else
        cmdPause.Caption = "Resume"
        pause = True
    End If
End Sub

Private Sub cmdStop_Click()
     stopIt = True
End Sub

Private Sub Command2_Click()
    ShellnWait "notepad """ & LastList & """", vbNormalFocus
    Globals.Initalize True 'reads in changes to testscript
End Sub

Private Sub List1_DblClick()
    i = List1.ListIndex
    ans = InputBox("If you want to set a default value for this form element fill it in below. If not leave empty")
    If ans = Empty Then Exit Sub
    With wb.Document.Forms(formIndex)
        List1.List(i) = "DEFAULT - " & .elements(i).Name & "=" & ans
    End With
End Sub

Private Sub Form_Load()
    Call Globals.Initalize
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Height > 5888 Then wb.Height = Me.Height - wb.Top - 450
  If Me.Width > 7035 Then
    wb.Width = Me.Width - wb.Left - 200
    List1.Width = Me.Width - List1.Left - 200
  End If
End Sub

Private Sub form_unload(cancel As Integer)
    Unload Me
End Sub

Private Sub tmrDelay_Timer()
  If Not pause Then WaitingToReturn = False
End Sub

Private Sub tmrTimeout_Timer()
   If Not pause Then
        timeOutOveride = True
        tmrTimeout.Enabled = False
   End If
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, cancel As Boolean)
    tmrTimeout.Enabled = True
    timeOutOveride = False
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, url As Variant)
    If WaitingToReturn Then WaitingToReturn = False
End Sub

Private Sub wb_NewWindow2(ppDisp As Object, cancel As Boolean)
    cancel = True
End Sub


