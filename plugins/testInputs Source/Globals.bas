Attribute VB_Name = "Globals"
Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Global frmMain As Object
Global frmRawRequest As Object
Global frmAnalyze As Object

Global testString()
Global LastList As String
Private Initalized As Boolean

Sub Initalize(Optional force As Boolean = False)
    If Initalized And Not force Then Exit Sub
    LastList = App.path & "\LastList.txt"
    Erase testString()
    If FileExists(LastList) Then
        tmp = Split(ReadFile(LastList), vbCrLf)
        For i = 0 To UBound(tmp)
            If tmp(i) <> Empty Then push testString(), tmp(i)
        Next
    Else
        testString() = LoadDefaultTestStrings
    End If
    Initalized = True
End Sub

Sub ShellnWait(cmdline, focus As VbAppWinStyle)
 On Error GoTo oops
    pid = Shell(cmdline, focus)
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    GetExitCodeProcess hdlProg, lExitCode
    Do While lExitCode = STILL_ACTIVE
        DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
    CloseHandle hdlProg
 Exit Sub
oops: MsgBox "Err in Shellnwait with cmdline:" & vbCrLf & vbCrLf & cmdline & vbCrLf & vbCrLf & Err.Description
End Sub

Function LoadDefaultTestStrings()
    Dim tmp()
    push tmp(), "# Shows some of the basic language syntax click ? for more help"
    push tmp(), "INJECT '<h1><center>Html Inclusion' LOOKFOR IT LOG 'Allows Html Input'"
    push tmp(), "IF NOT 'Allows Html Input' THEN GOTO 1"
    push tmp(), "INJECT '<script></script>' LOOKFOR IT LOG 'Allows Script Blocks'"
    push tmp(), "INJECT '<img src=javascript:>' LOOKFOR IT LOG 'Allow Img Src Scripts'"
    push tmp(), "INJECT '<img onerror=vbscript:>' LOOKFOR IT LOG 'Allows Img Src Event Handlers'"
    push tmp(), "1"
    push tmp(), "INJECT "";'"" LOOKFOR SQLERROR LOG 'Possible Sql Command Insertion Point'"
    LoadDefaultTestStrings = tmp()
End Function

Function wasBadSqlError(it) As Boolean
    'this function may get very complex this is place holder!
    If AnyOfTheseInstr(it, "Database error #:,SQL Server error,SQL Error,ODBC Error,SQL Syntax") _
    And NotTheseInstr(it, "Conversion to,Clng,Cbool") Then
        wasBadSqlError = True
    Else
        wasBadSqlError = False
    End If
End Function

Function AnyOfTheseInstr(s, them) As Boolean
    t = Split(them, ",")
    For i = 0 To UBound(t)
        If InStr(1, s, t(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True: Exit Function
        End If
    Next
End Function

Function NotTheseInstr(s, them) As Boolean
    t = Split(them, ",")
    For i = 0 To UBound(t)
        If InStr(1, s, t(i), vbTextCompare) > 0 Then
            NotTheseInstr = False: Exit Function
        End If
    Next
    NotTheseInstr = True
End Function


'------------------------------------------------------------------------
'Library Code
'------------------------------------------------------------------------
Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function ReadFile(filename)
  f = FreeFile
  Temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = Temp
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

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    ReDim Preserve ary(UBound(ary) + 1) '<-throws Error If Not initalized
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
    On Error GoTo oops
    X = UBound(ary)
    AryIsEmpty = False
    Exit Function
oops: AryIsEmpty = True
End Function

Function CountOccurances(it, find) As Integer
    If InStr(it, find) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, find)
    CountOccurances = UBound(tmp)
End Function


