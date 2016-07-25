Attribute VB_Name = "General"
Global blockServers() As String      'list of servers we can auto cancel navigation to (ad servers)
Global Bookmarks() As String         'ughh our ughh huhuhu bookmarks :P
Global ActionLogFile As String       'path to log file when in log mode
Global BookmarkFile As String        'path to bookmarks file
Global BlockServerFile As String     'path to what block servers list
Global NotesFile As String           'path to our default notes file
Global EmbedableScriptEnv As String  'path to Included script file
Global ProxyListFile As String       'path to our list of proxies

Sub SetGlobalFilePaths()

    Dim BasePath As String
    
    p1 = App.path & "\config"
    p2 = App.path & "\..\config"
    
    If FolderExists(p1) Then
        BasePath = p1 & "\"
    ElseIf FolderExists(p2) Then
        BasePath = p2 & "\"
    Else
        MkDir IIf(FolderExists(App.path & "\source"), p1, p2)
        Call SetGlobalFilePaths
        Exit Sub
    End If
    
    ProxyListFile = BasePath & "ProxyList.txt"
    
    NotesFile = BasePath & "Notes.txt"
    If FileExists(NotesFile) Then frmMain.rtfNotes.text = ReadFile(NotesFile)
        
    EmbedableScriptEnv = BasePath & "EmbedableScriptEnv.html"
    If Not FileExists(EmbedableScriptEnv) Then EmbedableScriptEnv = Empty
        
    BlockServerFile = BasePath & "BlockedServers.txt"
        
        If Not FileExists(BlockServerFile) Then
            blockServers() = Split("*doubleclick*,*fusion*,*ad*.com*,*Ads.asp*", ",")
        Else
            blockServers() = Split(ReadFile(BlockServerFile), vbCrLf)
        End If
        
        frmMain.txtBlockServers = Join(blockServers, vbCrLf)
                          
        BookmarkFile = BasePath & "Bookmarks.txt"
        If FileExists(BookmarkFile) Then
            Dim tmp() As String
            tmp() = Split(ReadFile(BookmarkFile), vbCrLf)
            If Not aryIsEmpty(tmp) Then
                For i = 0 To UBound(tmp)
                    If tmp(i) <> Empty Then
                        push Bookmarks(), tmp(i)
                        Load frmMain.mnuBookmarkItem(i + 1)
                        frmMain.mnuBookmarkItem(i + 1).caption = ExtractKey(Bookmarks(i))
                    End If
                Next
            End If
        End If

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

Sub ShowRtClkMenu(f As Form, t As Object, m As Menu)
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

Function ExtractValue(s)
        On Error Resume Next
        ExtractValue = Mid(s, InStr(s, "=") + 1, Len(s))
End Function

Function ExtractKey(s)
        On Error Resume Next
        ExtractKey = Mid(s, 1, InStr(s, "=") - 1)
End Function


