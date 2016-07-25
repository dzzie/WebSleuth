VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Sub SetHost(newref As Object)
    'sleuth passes in frmMain object, we set it as globals in this
    'project and use functions from sleuth frmmain to get other 2 forms
    'unfornatly i cant find a way so far to use public module functions :(
    'although you can run public form functions.
    
    Set frmMain = newref
    Set frmAnalyze = frmMain.getfrmAnalyze
    Set frmRawRequest = frmMain.getfrmRawReq
End Sub

Function HookMenu() As String()
    'here is where we add to the sleuths menu..you must know what
    'form & menu name you want to add to and only add to a menu that supports
    'this or else your startup event will never be fired!
    
    'note you will be typing without intellisense so set up this page!
    'see bottom of this page for all the other places you can hook

    Dim ret() As String
    
    i = frmMain.mnuListPlugin.count
    Load frmMain.mnuListPlugin(i)
    frmMain.mnuListPlugin(i).Caption = "Test Inputs"
    frmMain.mnuListPlugin(i).Visible = True
    push ret(), "TestInputs.frmMain.mnuListPlugin." & i & ".A."
    
    i = frmRawRequest.mnuPlugins.count
    Load frmRawRequest.mnuPlugins(i)
    frmRawRequest.mnuPlugins(i).Caption = "Unix 2 Dos"
    frmRawRequest.mnuPlugins(i).Visible = True
    push ret(), "TestInputs.frmRawRequest.mnuPlugins." & i & ".B."
    
    i = frmMain.mnuPlugins.count
    Load frmMain.mnuPlugins(i)
    frmMain.mnuPlugins(i).Caption = "Http Put/Delete"
    frmMain.mnuPlugins(i).Visible = True
    push ret(), "TestInputs.frmMain.mnuPlugins." & i & ".C."


    HookMenu = ret()
    
    'return string= "dllname.frmMenuIsOn.mnuName.indexYouAdd.ArgumentYouRecieve."
    'ending dot is important! it is your responsibility that it is there!
End Function

Sub StartUp(arg)
    
    Select Case arg
        Case "A"
            If frmMain.Combo1.ListIndex = 4 Then 'forms are being shown
                 frmTestForm.InitalizeFromForm frmMain.wb.Document.location.href, frmMain.List.SelectedIndex
            Else
                 frmTestForm.InitalizeFromUrl frmMain.List.SelectedText
            End If
            frmTestForm.Show
        Case "B"
            If frmRawRequest.Text1.SelLength = 0 Then MsgBox "Nothing Selected!": Exit Sub
            frmRawRequest.Text1.SelText = UnixToDos(frmRawRequest.Text1.SelText)
        Case "C"
            frmPUT.Show
    End Select
    
End Sub

Private Sub push(ary() As String, value) 'this modifies parent ary object
    On Error GoTo init
    ReDim Preserve ary(UBound(ary) + 1) '<-throws Error If Not initalized
    ary(UBound(ary)) = CStr(value)
    Exit Sub
init: ReDim ary(0): ary(0) = CStr(value)
End Sub

Private Function UnixToDos(it) As String
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

'Rember you can also access all the object and public functions on
'these forms as well !
'
'  Right click menu in textbox of frmAnalyze
'    i = frmAnalyze.mnuPlugins.Count
'    Load frmAnalyze.mnuPlugins(i)
'    frmAnalyze.mnuPlugins(i).Caption = "TestInputs"
'    frmAnalyze.mnuPlugins(i).Visible = True
'    push ret(), "TestInputs.frmAnalyze.mnuPlugins." & i & ".A."

'  Main Web Sleuth Menu "Rt ClickMenu"
'    i = frmMain.mnuPlugins.Count
'    Load frmMain.mnuPlugins(i)
'    frmMain.mnuPlugins(i).Caption = "TestInputs"
'    frmMain.mnuPlugins(i).Visible = True
'    push ret(), "TestInputs.frmMain.mnuPlugins." & i & ".A."

'  Raw Http Request textbox right click menu
'    i = frmRawRequest.mnuPlugins.Count
'    Load frmRawRequest.mnuPlugins(i)
'    frmRawRequest.mnuPlugins(i).Caption = "TestInputs"
'    frmRawRequest.mnuPlugins(i).Visible = True
'    push ret(), "TestInputs.frmRawRequest.mnuPlugins." & i & ".A."
