Attribute VB_Name = "SleuthFx"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Global Plugin() As Object
Global RegisteredPlugins() As String

Global objSelection As Object
Global objSelDocument As HTMLDocument

'--------------------------------------------------------------------
'Main Entry Point into Program
'--------------------------------------------------------------------
Sub Main()
    Load frmMain
    LoadPlugins
    frmMain.Show
    frmRawRequest.Hide
    ' so wacky...the testinputs plugin adds a menu item to
    ' this form, that is enough make the form load and show
    ' i guess that makes sense but was weird to find
End Sub


'-------------------------------------------------------------------------
'PlugIn Code
'-------------------------------------------------------------------------
Sub LoadPlugins()
    If Not CBool(frmMain.chkLoadPlugins.value) Then Exit Sub
    
    Dim tmp() As String
    ReDim Plugin(0)
    BasePath = App.path & "\Plugins"
    If Not FolderExists(BasePath) Then BasePath = App.path & "\..\Plugins"
    tmp = GetFolderFiles(BasePath, "*dll")
    On Error GoTo oops
    If aryIsEmpty(tmp) Then Exit Sub
    For i = 0 To UBound(tmp)
        tmp(i) = GetBaseName(tmp(i))
        ReDim Preserve Plugin(i)
        Set Plugin(i) = CreateObject(tmp(i) & ".plugin")
        Plugin(i).SetHost frmMain
        'format of reg plugins= frmName.mnuName.mnuindex.StartArgument.pluginIndex
        Dim ret() As String
        ret() = Plugin(i).HookMenu
        For j = 0 To UBound(ret)
          push RegisteredPlugins(), ret(j) & i
        Next
    Next
    Exit Sub
oops:
        'err 429 = failed to create obj
        If Err.Number = 429 Then
           If trytoRegister(tmp(i)) = True Then
                MsgBox "Plugin Registered Successfully!", vbInformation
                Set Plugin(i) = CreateObject(tmp(i) & ".plugin")
           Else
               MsgBox "Failed to Register: " & tmp(i) & " make sure only websleuth plugins are in the plugins folder!"
           End If
        Else
            MsgBox "Failed to load plugin " & tmp(i) & vbCrLf & vbCrLf & Err.Number & ":" & Err.Description
        End If
        Resume Next
End Sub

Private Function trytoRegister(dllName) As Boolean
    If MsgBox("New Plugin Found Would you like to register " & dllName & ".dll?" & vbCrLf & vbCrLf & "Note that running plugins from untrusted parties is the same as running an untrusted exe!", vbYesNo + vbExclamation) = vbYes Then
        BasePath = App.path & "\Plugins"
        If Not FolderExists(BasePath) Then BasePath = App.path & "\..\Plugins"
        dllpath = BasePath & "\" & dllName & ".dll"
        If FileExists(dllpath) Then
           trytoRegister = RegisterServer(CStr(dllpath))
        Else
            MsgBox "OOps i got the path wrong filenot found!"
        End If
    End If
End Function

Private Function RegisterServer(DllServerPath As String, Optional unRegister = False) As Boolean
    On Error Resume Next
    lb = LoadLibrary(DllServerPath)
    
    If unRegister Then pa = GetProcAddress(lb, "DllUnregisterServer") _
    Else pa = GetProcAddress(lb, "DllRegisterServer")
    
    If CallWindowProc(pa, frmMain.hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then RegisterServer = True
    
    FreeLibrary lb
End Function

Sub FirePluginEvent(index As Integer, connectionString)
    On Error GoTo oops
    For i = 0 To UBound(RegisteredPlugins)
        tmp = Split(RegisteredPlugins(i), ".")
        '3=mnuIndex, 4=StartArgument, 5=ObjectIndex
        If InStr(1, RegisteredPlugins(i), connectionString, vbTextCompare) > 0 _
         And tmp(3) = index Then
            Call Plugin(tmp(5)).startUp(tmp(4))
            Exit For
        End If
    Next
    Exit Sub
oops: MsgBox "Plugin failed!, Plugin String: " & RegisteredPlugins(i) & vbCrLf & "Error: " & Err.Description
End Sub

'---------------------------------------------------------------------
'Sleuth specific document parsing routines below
'---------------------------------------------------------------------
Function ExtractPostData(it) As String 'Code submitted by: dhurst@spidynamics.com
        lLen = LenB(it)                'Use LenB to get the byte count
        Dim strPostData As String      'If it's a post form, lLen will be > 0
        If lLen > 0 Then               'Use MidB to get 1 byte at a time
          For lCount = 1 To lLen
              strPostData = strPostData & Chr(AscB(MidB(it, lCount, 1)))
          Next
        End If
        ExtractPostData = strPostData
End Function

Function GetLinks(d As HTMLDocument, Optional ary = Empty) 'As Variant()
        Dim tmp()
        For i = 0 To d.links.length - 1
            it = d.links(i).href
            If IsArray(ary) Then push ary, "Link - " & it _
            Else push tmp(), it
        Next
        If aryIsEmpty(tmp) Then push tmp, "No Links In Document"
        GetLinks = IIf(IsArray(ary), ary, tmp())
End Function

Function GetImages(d As HTMLDocument, Optional ary = Empty) 'As Variant()
        Dim tmp()
        For i = 0 To d.images.length - 1
            it = d.images(i).src
            If IsArray(ary) Then push ary, "Image - " & it _
            Else push tmp(), it
        Next
        If aryIsEmpty(tmp) Then push tmp, "No Images In Document"
        GetImages = IIf(IsArray(ary), ary, tmp())
End Function

Function GetFrames(d As HTMLDocument, Optional ary = Empty) 'As Variant()
        On Error GoTo oops
        Dim tmp()
                
        For i = 0 To d.frames.length - 1
             it = d.frames(i).Name & " - " & d.frames(i).Document.location.href
             If IsArray(ary) Then push ary, "Frame - " & it _
             Else push tmp(), it
        Next
        
        If aryIsEmpty(tmp) Then push tmp(), "No Frames In Document"
        GetFrames = IIf(IsArray(ary), ary, tmp())
Exit Function
oops:
        it = Replace(Err.Description, vbCrLf, Empty)
        If IsArray(ary) Then push ary, it Else: push tmp(), it
        Resume Next
End Function

Function GetForms(d As HTMLDocument)
        Dim tmp()
        For i = 0 To d.Forms.length - 1
           With d.Forms(i)
              push tmp(), UCase(.method) & " - " & .Name & " - " & "  " & .Action
           End With
        Next
        If aryIsEmpty(tmp) Then push tmp(), "No Forms In Document"
        GetForms = tmp()
End Function

Function GetScripts(d As HTMLDocument, Optional ary = Empty)
        Dim ret()
        For i = 0 To d.scripts.length - 1
          With d.scripts(i)
            it = .src
            If it = Empty Then
                it = d.scripts(i).innerHTML
            Else
                it = "SRC = " & it
            End If
            If IsArray(ary) Then push ary, "Script - " & it _
            Else push ret(), it
          End With
        Next
        If aryIsEmpty(ret) Then push ret(), "No Scripts In Document"
        If IsArray(ary) Then GetScripts = ary Else GetScripts = ret()
End Function

Function GetComments(d As HTMLDocument)
   On Error Resume Next
        Dim tmp()
        doc = ParseScript(d.body.innerHTML)
        c = Split(doc, "<!")
        If UBound(c) > 0 Then
            For i = 0 To UBound(c)
                push tmp, "<!" & Mid(c(i), 1, InStr(1, c(i), ">"))
            Next
        End If
        If aryIsEmpty(tmp) Then push tmp, "No Comments in Document"
        GetComments = tmp()
End Function

Function GetFormsContent(d As HTMLDocument, index)
    Dim ret()
    With d.Forms(index)
        push ret(), "Form: " & .Name & " Method:" & UCase(.method)
        push ret(), "ACTION: " & .Action
        push ret(), "BASE URL: " & d.location.href
        For j = 0 To .elements.length - 1
            push ret(), UCase(.elements(j).Type) & " - " & .elements(j).Name & "=" & .elements(j).value
        Next
    End With
    GetFormsContent = ret
End Function

Function GetPageStats(d As HTMLDocument, ret, Optional IncludeSource As Boolean = False)
    push ret, String(75, "-") & vbCrLf
    push ret, vbCrLf & vbCrLf & "Page: " & d.location.href
    push ret, "Cookie: " & IIf(d.cookie = "", "No Cookies In Document", d.cookie) & vbCrLf
    push ret, "Links: " & vbCrLf & vbTab & Join(GetLinks(d), vbCrLf & vbTab)
    push ret, "Images: " & vbCrLf & vbTab & Join(GetImages(d), vbCrLf & vbTab)
    push ret, "Scripts: " & vbCrLf & vbTab & Join(GetScripts(d), vbCrLf & vbTab)
    push ret, "Comments: " & vbCrLf & vbTab & Join(GetComments(d), vbCrLf & vbTab)
    push ret, "MetaTags: " & vbCrLf & vbTab & Join(GetMetaTags(d), vbCrLf & vbTab)
    
    push ret, "Forms: " & vbCrLf & vbTab & Join(GetForms(d), vbCrLf & vbTab)
    For i = 0 To d.Forms.length - 1
        push ret, vbTab & Join(GetFormsContent(d, i), vbCrLf & vbTab & vbTab)
    Next
    
    If IncludeSource Then
        push ret, vbCrLf & String(75, "#") & vbCrLf & "body.innerHTML" & vbCrLf
        push ret, d.body.innerHTML & vbCrLf & String(75, "#") & vbCrLf
    End If
    
    push ret, "Frames: " & vbCrLf & vbTab & Join(GetFrames(d), vbCrLf & vbTab)
    For i = 0 To d.frames.length - 1
        On Error Resume Next
        push ret, Join(GetPageStats(d.frames(i).Document, ret, IncludeSource), vbCrLf & vbTab & vbTab & vbTab)
    Next
End Function

Function GetFormHTML(d As HTMLDocument, index)
    tmp = Split(d.body.innerHTML, "<form", , vbTextCompare)
    it = tmp(index + 1)
    it = "<form" & Mid(it, 1, InStr(5, it, "</form>", vbTextCompare) + 6)
    GetFormHTML = ParseScript(it)
End Function

Function GetMetaTags(d As HTMLDocument)
    Dim ret()
    tmp = Split(d.body.innerHTML, "<meta", , vbTextCompare)
    For i = 1 To UBound(tmp)
        it = tmp(i)
        push ret, "<meta" & Mid(it, 1, InStr(1, it, ">", vbTextCompare) + 1)
    Next
    If aryIsEmpty(ret) Then push ret, "No Meta Tags in Document"
    GetMetaTags = ret
End Function

Function BreakDownCookie(it)
    Dim ret()
    If it = Empty Then
        push ret(), "No Cookies In Document"
    Else
        tmp = Split(it, "&")
        For i = 0 To UBound(tmp)
            If InStr(tmp(i), ";") Then
                t = Split(tmp(i), ";")
                For j = 0 To UBound(t)
                    push ret(), LTrim(t(j))
                Next
            Else
                push ret(), LTrim(tmp(i))
            End If
         Next
     End If
     BreakDownCookie = ret()
End Function


Sub HTMLTransform(d As HTMLDocument, h As Integer)
   ' On Error GoTo oops
    
    tmp = d.body.innerHTML
    Select Case h
        Case 0: 'EmbedScript
            If Not FileExists(EmbedableScriptEnv) Then
                MsgBox EmbedableScriptEnv & " not found!", vbCritical
                Exit Sub
            End If
            tmp = tmp & "<HR>" & ReadFile(EmbedableScriptEnv) & " "
        Case 1: 'Hidden2Text
            tmp = Replace(tmp, "type=hidden", "type=text", , , vbTextCompare) & " "
        Case 2: 'Select2Text
            tmp = Replace(tmp, "</select>", Empty, , , vbTextCompare)
            tmp = Replace(tmp, "<option", "&lt;option", , , vbTextCompare)
            tmp = Replace(tmp, "<select", "<input type=text", , , vbTextCompare)
        Case 3: 'Check2Text
            tmp = Replace(tmp, "type=check", "type=text ", , , vbTextCompare)
        Case 4: 'Radio2Text
            MsgBox "Soon"
        Case 5: 'BreakScripts
            tmp = ParseScript(tmp)
    End Select
    
        d.body.innerHTML = tmp & " "
        frmMain.wb.Stop
Exit Sub
oops: MsgBox "Err in HTML transform: " & Err.Description, vbCritical
End Sub

Sub GenReport()
    Dim ret(), IncludeSource As Boolean
    IncludeSource = IIf(MsgBox("Do you want to include each docs body.innerHTML ?", vbQuestion + vbYesNo) = vbYes, True, False)
    fpath = App.path & "\Sleuth_Report.txt"
    
    push ret(), vbCrLf & String(75, "-")
    push ret(), Date & String(5, " ") & Time & String(5, " ") & "Saved as: " & fpath & vbCrLf
    push ret(), "If you want to save this file be sure to do a SAVE AS or"
    push ret(), "else it will be automatically overwritten by next report!" & vbCrLf
    
    Call GetPageStats(frmMain.wb.Document, ret, IncludeSource)
    WriteFile fpath, Join(ret, vbCrLf)
    Shell "notepad """ & fpath & """", vbNormalFocus
    
End Sub

Function GetSelectedHtml(d As HTMLDocument) As String
On Error GoTo oops
    'Dim k As Object
    'Set k = d.selection.createRange
    'GetSelectedHtml = CStr(k.htmlText)
    
    Set objSelDocument = d
    Set objSelection = d.selection.createRange
    GetSelectedHtml = CStr(objSelection.htmlText)

Exit Function
oops:
MsgBox Err.Description, vbExclamation, "GetSelectedHtml"
End Function



Function TurnFormIntoQueryString(d As HTMLDocument, formIndex As Integer) As String
    Dim url As String
    Dim ret() As String
    
    With d.Forms(formIndex)
        q = InStr(.Action, "?")
        If q > 0 Then
            If .method = "POST" Then
                MsgBox "This form action is sent with querystring arguments...because of the way this works now they will all be grouped with other form values..this may cause problems because this form is a POST and they will be in thePOST body now sorry"
            End If
        End If
        
        If InStr(.Action, "http://") < 1 Then
            If InStr(8, d.location, "/") Then
                'remove any querystring args from current page
                base = Mid(d.location, 1, InStrRev(d.location, "/"))
            Else
                base = d.location
            End If
            url = Replace(base, .Action, Empty) & "/" & .Action & IIf(q > 0, "&", "?")
            MsgBox "Your going to have to verify path is right", vbInformation
        Else
            url = .Action & IIf(q > 0, "&", "?")
        End If
        
        For j = 0 To .elements.length - 1
            url = url & .elements(j).Name & "=" & .elements(j).value & "&"
        Next
    End With
    
    'this sucks
    url = Replace(url, "://", ":/:")
    While InStr(url, "//") > 0
        url = Replace(url, "//", "/")
    Wend
    url = Replace(url, ":/:", "://")
    
    TurnFormIntoQueryString = url
    
    'MsgBox url
    
End Function












