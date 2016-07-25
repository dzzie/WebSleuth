Attribute VB_Name = "StdFx"
Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Global Plugin() As Object
Global RegisteredPlugins() As String

'--------------------------------------------------------------------
'Main Entry Point into Program
'--------------------------------------------------------------------
Sub Main()
    Load frmMain
    LoadPlugins
    frmMain.Show
End Sub


'-------------------------------------------------------------------------
'PlugIn Code
'-------------------------------------------------------------------------
Private Sub LoadPlugins()
    Dim tmp() As String
    ReDim Plugin(0)
    tmp = GetFolderFiles(App.path & "\Plugins", "*dll")
    'tmp = GetFolderFiles("D:\Projects\web sleuth\Plugins", "*dll")
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
        dllpath = App.path & "\Plugins\" & dllName & ".dll"
        'dllpath = "D:\Projects\web sleuth\Plugins\" & dllName & ".dll"
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


'--------------------------------------------------------------------
' File system and path parsing functions
'--------------------------------------------------------------------
Function GetBaseName(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function FolderExists(path) As Boolean
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function SafeFileName(proposed) As String
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

Function GetFolderFiles(folder, Optional filter = ".*", Optional retFullPath As Boolean = True) As String()
   Dim fnames() As String
   
   If Not FolderExists(folder) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        'MsgBox "hey no folder!"
        Exit Function
   End If
   
   folder = IIf(Right(folder, 1) = "\", folder, folder & "\")
   If Left(filter, 1) = "*" Then extension = Mid(filter, 2, Len(filter))
   If Left(filter, 1) <> "." Then filter = "." & filter
   
   fs = Dir(folder & "*" & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folder & fs, fs)
     fs = Dir()
   Wend
   
   GetFolderFiles = fnames()
End Function

Function WebParentFolderFromURL(url) As String
    If url = Empty Or InStr(url, "/") < 1 Then Exit Function
    tmp = Split(url, "/")
    If InStr(tmp(UBound(tmp)), ".") > 0 Then tmp(UBound(tmp)) = Empty
    tmp = Join(tmp, "/")
    If Right(tmp, 2) = "//" Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    WebParentFolderFromURL = CStr(tmp)
End Function

Function WebFileNameFromPath(fullpath)
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = tmp(UBound(tmp))
    End If
End Function

Sub UseIE_DDE_GETCurrentPageURL(t As TextBox)
  On Error GoTo nope
    t.LinkTopic = "iexplore|WWW_GetWindowInfo"
    t.LinkItem = &HFFFFFFFF
    t.LinkMode = 2
    t.LinkRequest
    t.LinkTopic = Empty
    tmp = Split(t, ",")
    t = Replace(tmp(0), """", Empty)
  Exit Sub
nope: t.text = "about:blank"
End Sub

'-------------------------------------------------------------------
' general library functions
'-------------------------------------------------------------------
Function filt(txt, Remove As String)
  If Right(txt, 1) = "," Then txt = Mid(txt, 1, Len(txt) - 1)
  tmp = Split(Remove, ",")
  For i = 0 To UBound(tmp)
     txt = Replace(txt, tmp(i), "", , , vbTextCompare)
  Next
  filt = txt
End Function

Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    ReDim Preserve ary(UBound(ary) + 1) '<-throws Error If Not initalized
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function UnixToDos(it) As String
    If InStr(it, vbLf) > 0 Then
        tmp = Split(it, vbLf)
        For i = 0 To UBound(tmp)
            If InStr(tmp(i), vbCr) < 1 Then tmp(i) = tmp(i) & vbCr
        Next
        UnixToDos = Join(tmp, vbLf)
    Else
        UnixToDos = CStr(it)
    End If
End Function

Function BatchReplace(ByVal it, them, Optional compare As VbCompareMethod = vbTextCompare) As String
    'it=data string, them="changeThis->toThis,andThis->toThat"
    t = Split(them, ",")
    For i = 0 To UBound(t)
        If InStr(t(i), "->") > 1 Then
            s = Split(t(i), "->")
            it = Replace(it, s(0), s(1), , , compare)
        End If
    Next
    BatchReplace = CStr(it)
End Function

Sub ShowRtClkMenu(f As Form, t As TextBox, m As Menu)
        LockWindowUpdate t.hWnd
        t.Enabled = False
        DoEvents
        f.PopupMenu m
        t.Enabled = True
        LockWindowUpdate 0&
End Sub

Function aryIsEmpty(ary) As Boolean
    On Error GoTo oops
    x = UBound(ary)
    aryIsEmpty = False
    Exit Function
oops: aryIsEmpty = True
End Function

Function CountOccurances(it, find) As Integer
    If InStr(1, it, find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, find, , vbTextCompare)
    CountOccurances = UBound(tmp)
End Function

Sub WipeStrAry(ary)
    Dim tmp() As String
    ary = tmp
End Sub

Function GetPathsInStep(fullpath) As String()
    Dim ret() As String
    'fullpath = var/www/htdocs/
    'ret(0) = fullpath
    'ret(1) = var/
    'ret(2) = var/www/
    'ret(3) = var/www/htdocs/
    push ret(), fullpath
    tmp = Split(fullpath, "/")
    Dim it
    For i = 0 To UBound(tmp)
        it = it & IIf(it = Empty, Empty, "/") & tmp(i)
        push ret(), it
    Next
    GetPathsInStep = ret()
End Function

Sub pop(ary, Optional Count = 1) 'this modifies parent ary obj
    If Count > UBound(ary) Then ReDim ary(0)
    For i = 1 To Count
        ReDim Preserve ary(UBound(ary) - 1)
    Next
End Sub

Function ExtractValue(s)
        On Error Resume Next
        ExtractValue = Mid(s, InStr(s, "=") + 1, Len(s))
End Function

Function ExtractKey(s)
        On Error Resume Next
        ExtractKey = Mid(s, 1, InStr(s, "=") - 1)
End Function

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
            If it = Empty Then it = "Embeded Script"
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
        For i = 0 To UBound(c)
            push tmp, "<!" & Mid(c(i), 1, InStr(1, c(i), ">"))
        Next
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
    push ret, "Cookie: " & IIf(d.Cookie = "", "No Cookies In Document", d.Cookie) & vbCrLf
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

Function GetScriptContent(d As HTMLDocument, index)
    On Error GoTo oops
    tmp = Split(d.body.innerHTML, "<script", , vbTextCompare)
    Dim ret()
    For i = 1 To UBound(tmp)
        s = InStr(1, tmp(i), "</script>", 1)
        If s < 1 Then s = Len(tmp(i))
        push ret(), "<script " & Mid(tmp(i), 1, s + 9) & vbCrLf
    Next
    GetScriptContent = Join(ret, vbCrLf)
    Exit Function
oops:
    l = Len(d.body.innerHTML)
    Msg = "Page probably has scripts in <head> section which throws off index for all scripts!"
    If l = 0 Then Msg = " Page Content Length was 0, this occurs when page author does not include a <body> tag."
    MsgBox "Oops there was an error extracting script," & vbCrLf & vbCrLf & Msg, vbExclamation
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
            If Not FileExists(frmMain.EmbedableScriptEnv) Then
                MsgBox frmMain.EmbedableScriptEnv & " not found!", vbCritical
                Exit Sub
            End If
            tmp = tmp & "<HR>" & ReadFile(frmMain.EmbedableScriptEnv) & " "
        Case 1: 'Hidden2Text
            tmp = Replace(tmp, "type=hidden", "type=text", , , vbTextCompare) & " "
        Case 2: 'Select2Text
            tmp = Replace(tmp, "</select>", Empty, , , vbTextCompare)
            tmp = Replace(tmp, "<option", "&lt;option", , , vbTextCompare)
            tmp = Replace(tmp, "<select", "<input type=text", , , vbTextCompare)
        Case 3: 'Check2Text
            tmp = Replace(tmp, "type=check", "type=text ", , , vbTextCompare)
        Case 4: 'Radio2Text
        
        Case 5: 'BreakScripts
            tmp = ParseScript(tmp)
    End Select
    
        d.body.innerHTML = tmp & " "
        frmMain.wb.Stop
Exit Sub
oops: MsgBox "Err in HTML transform: " & Err.Description, vbCritical
End Sub
