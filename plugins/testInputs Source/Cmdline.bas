Attribute VB_Name = "Cmdline"
'Info:     this function is used for grabbing the command line
'           info passed in from a text string. it will recgonize
'           double and single quoted arguments with spaces as one argument
'           it will also recgonize " -ac" as two switchs
'
'License:  you are free to use this library in your personal projects, so
'               long as this header remains inplace. This code cannot be
'               used in any project that is to be sold. This source code
'               can be freely distributed so long as this header reamins
'               intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie


Function GetArgs(cmd) As Variant()

    If cmd = Empty Then Exit Function
    
    Dim args()
    
    Dim inminus As Boolean
    Dim isword As Boolean
    
    tmp = ""
    lastLetter = ""
    
    For i = 1 To Len(cmd)
      letter = Mid(cmd, i, 1)
      nextlet = Mid(cmd, i + 1, 1)
      
      Select Case letter
        Case "-":
                  If lastLetter = " " Then
                        inminus = True: isword = False
                  End If
        Case " ":
                    inminus = False
                    If isword Then
                      isword = False
                      push args, tmp
                      tmp = ""
                    End If
        Case "'":
                   isword = False
                   x = InStr(i + 1, cmd, "'")
                   push args, Mid(cmd, i + 1, x - i)
                   i = x + 1
                   GoTo nextcycle
        Case """":
                   isword = False
                   x = InStr(i + 1, cmd, """")
                   push args, Mid(cmd, i + 1, x - i)
                   i = x + 1
                   GoTo nextcycle
      End Select
      
      If inminus And letter <> "-" Then
         push args, letter
      Else
         isword = True
         tmp = tmp & letter
         If i = Len(cmd) Then push args, tmp
      End If
      lastLetter = letter
      
nextcycle:
    Next
    
    If AryIsEmpty(args) Then Exit Function
    
    For i = 0 To UBound(args)
        args(i) = Trim(LTrim(args(i)))
    Next
    
    If args(UBound(args)) = "'" Or args(UBound(args)) = Empty Then pop args
    
    GetArgs = args()
End Function

Private Sub pop(ary, Optional count = 1) 'this modifies parent ary obj
    If count > UBound(ary) Then ReDim ary(0)
    For i = 1 To count
        ReDim Preserve ary(UBound(ary) - 1)
    Next
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
