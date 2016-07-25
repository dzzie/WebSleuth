VERSION 5.00
Begin VB.UserControl List 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   ScaleHeight     =   1080
   ScaleWidth      =   3555
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
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
      Height          =   750
      Left            =   105
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   105
      Width           =   3165
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event RightClick()
Event Click()
Event DoubleClick()

'Property Let MultiSelect(v As Boolean)
'  lst.MultiSelect = v
'End Property

Property Get Count() As Long
  Count = lst.ListCount
End Property
 
Property Get Value(index As Long)
    Value = lst.List(index)
End Property
 
Property Get SelectedText() As String
    SelectedText = lst.List(lst.ListIndex)
End Property

Property Get SelectedIndex() As Long
    SelectedIndex = lst.ListIndex
End Property

Sub Remove(index)
    lst.RemoveItem index
End Sub

Sub Clear()
    lst.Clear
End Sub

Sub AddItem(it)
    lst.AddItem it
End Sub

Sub UpdateValue(newVal, index)
    lst.List(index) = newVal
End Sub

Sub FilterList(filt, Optional Likeit As Boolean = True)
    If filt = "*" Or Trim(filt) = Empty Then Exit Sub
    Dim tmp()
    tmp() = GetListToArray
    tmp() = filterArray(tmp(), filt, Likeit)
    LoadArray tmp()
End Sub

Function GetListToArray() As Variant()
    Dim tmp()
    For i = 0 To lst.ListCount - 1
        push tmp, lst.List(i)
    Next
    GetListToArray = tmp()
End Function

Function GetListContents(Optional JoinWith = vbCrLf) As String
    Dim tmp As String
    For i = 0 To lst.ListCount - 1
        tmp = tmp & lst.List(i) & JoinWith
    Next
    GetListContents = tmp
End Function

Sub LoadFile(fpath, Optional delimiter = vbCrLf, Optional AppendIt As Boolean = False)
    If FileExists(fpath) Then
        tmp = Split(ReadFile(fpath), delimiter)
        LoadArray tmp, AppendIt
    End If
End Sub

Sub LoadArray(ary, Optional AppendIt As Boolean = False)
    If Not AppendIt Then lst.Clear
    If aryIsEmpty(ary) Then: lst.AddItem "[Empty Set]": Exit Sub
    For i = LBound(ary) To UBound(ary)
        lst.AddItem ary(i)
    Next
End Sub

Sub LoadDelimitedString(dStr, delimiter, Optional AppendIt As Boolean = False)
    tmp = Split(dStr, delimiter)
    LoadArray tmp, AppendIt
End Sub

Function filterArray(ary, filtStr, Optional Likeit As Boolean = True) As Variant()
    If aryIsEmpty(ary) Then Exit Function
    
    Dim tmp()
    filtStr = filtStr
    'if you use lcase() on somthing not expliticly defined string
    'it returns nothign! wildcard expression always second
    For i = LBound(ary) To UBound(ary)
        If Likeit Then
            If ary(i) Like filtStr Then push tmp, ary(i)
        Else
            If Not ary(i) Like filtStr Then push tmp, ary(i)
        End If
    Next
    
    filterArray = tmp()
End Function

Private Sub lst_Click()
    RaiseEvent Click
End Sub

Private Sub lst_DblClick()
    RaiseEvent DoubleClick
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RaiseEvent RightClick
End Sub

Private Sub lst_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FileExists(Data.Files(1)) Then LoadFile Data.Files(1)
End Sub

Private Sub UserControl_Initialize()
    lst.Top = 0
    lst.Left = 0
End Sub

Private Sub UserControl_Resize()
   lst.Height = UserControl.Height
   lst.Width = UserControl.Width
   'because lists only allow heights on vertain increments
   UserControl.Height = lst.Height
   UserControl.Width = lst.Width
End Sub

Sub MatchSize(it As Object)
    UserControl.Size it.Width, it.Height
End Sub


Private Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Private Sub AppendFile(path, it)
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Private Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init: ReDim ary(0): ary(0) = Value
End Sub

Private Function aryIsEmpty(ary) As Boolean
  On Error GoTo oops
    X = UBound(ary)
    aryIsEmpty = False
  Exit Function
oops: aryIsEmpty = True
End Function
