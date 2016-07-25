Attribute VB_Name = "parse"
Private Type lib 'b64 conversion library
  b64Chr(65) As String
  binHex(16) As String
  hexChr(16) As String
  Init As Boolean
End Type

Private lib As lib
Private escAry() 'characters always escaped
Private escAr2() 'characters not always escaped

'tries to keep somthing like format of document
Public Function parseHtml(info) As String
     Dim temp As String, EndOfTag As Integer
     fmat = Replace(info, "&nbsp;", " ")
     cut = Split(fmat, "<")

   For i = 0 To UBound(cut)
     EndOfTag = InStr(1, cut(i), ">")
        If EndOfTag > 0 Then
          EndOfText = Len(cut(i))
          NL = False
          If Left(cut(i), 2) = "br" Then NL = True
          cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
          If NL Then cut(i) = vbCrLf & cut(i)
          If cut(i) = vbCrLf Then cut(i) = ""
        End If
     temp = temp & cut(i)
    Next
    
    parseHtml = temp
End Function

Function ParseScript(info)
  Dim trimpage
  info = filt(info, "javascript,vbscript,mocha,createobject,activex,onclick,onmouse,onscroll,onkey,onload")
  scr = Split(info, "<script", , vbTextCompare)
  
  If aryIsEmpty(scr) Then ParseScript = info: Exit Function _
  Else: trimpage = scr(0)
  
  For i = 1 To UBound(scr)
    EndOfScript = InStr(1, scr(i), "</script>", vbTextCompare)
    trimpage = trimpage & Mid(scr(i), EndOfScript + 10, Len(scr(i)))
  Next
  
  ParseScript = trimpage
End Function

Function UnEscape(it)
    Dim f(): Dim c()
    n = Replace(it, "+", " ")
    If InStr(n, "%") > 0 Then
        t = Split(n, "%")
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), "%" & a
                push c(), b
            End If
        Next
        For i = 0 To UBound(f)
            n = Replace(n, f(i), c(i))
        Next
    End If
    UnEscape = n
End Function

Function escape(it, Optional fullEscape As Boolean = True)
    If aryIsEmpty(escAry) Then LoadEscapeArray
    If fullEscape Then
        For i = 0 To UBound(escAr2)
           h = Hex(escAr2(i))
           it = Replace(it, Chr(escAr2(i)), "%" & IIf(Len(h) = 1, "0" & h, h))
        Next
    End If
    For i = 0 To UBound(escAry)
       h = Hex(escAry(i))
       it = Replace(it, Chr(escAry(i)), "%" & IIf(Len(h) = 1, "0" & h, h))
    Next
    escape = it
End Function

Private Sub LoadEscapeArray()
    For i = 0 To 255
        If i > 47 And i < 58 Then GoTo skip  '1-10
        If i > 64 And i < 91 Then GoTo skip  'A-Z
        If i > 96 And i < 123 Then GoTo skip 'a-z
        Select Case i                      'not always escape
          Case 37, 35, 38, 46, 47, 58, 61, 63, 64: push escAr2(), i
          Case Else: push escAry(), i      'always have to escape
        End Select
skip:
    Next
End Sub



'------------------------------------------------------------------
'these are over a year old....I know they are slow but dont want to
'use any code i dont have explicit copyright on unless specifically
'donated by the author of the code...
'
'only use these on under 50k files or you will be waiting forever!
'------------------------------------------------------------------
Public Function b64Decode(it As String) As String
    Dim t() As String
    push t(), Empty
    it = Replace(Trim(it), vbCrLf, Empty)
    For i = 1 To Len(it)
        push t, Mid(it, i, 1)
    Next
    b64Decode = Join(b64DecodeEngine(t), Empty)
End Function

Public Function b64Encode(it As String) As String
    Dim t() As String
    push t(), Empty
    it = Replace(Trim(it), vbCrLf, Empty)
    For i = 1 To Len(it)
        push t, Mid(it, i, 1)
    Next
    b64Encode = Join(b64EncodeEngine(t), Empty)
End Function

Private Function b64EncodeEngine(it)
   On Error GoTo warn
    If Not lib.Init Then initAlpha
    
    Dim str As String  'it= BASE 1 string array of characters
    Dim s() As String  'returns BASE 1 string array of chars
    pad = 0            'how many times we had to pad val to encode (1-2)
    
    For i = 1 To UBound(it)  'ascii val-->hex val-->binary string
        it(i) = Hex2Bin(Hex(Asc(it(i))))
    Next
    
    str = Join(it, "")         'has to be div by 6 for now pad with 0's
    While Len(str) Mod 6 <> 0  'in final out put we must represent these as =
      str = str & "00"         'signs b64(64) which is 01000000 binary which cant
      pad = pad + 1            'be represented at this stage! so we must have counter :\
    Wend
    
    ReDim s(Len(str) Mod 6)    'fill s() with 6char div of str
    s = segment(str, 6)        'returns one based array!
                               
    For i = 1 To UBound(s)     'what letter corrosponds to it from
      s(i) = lib.b64Chr(Int("&H" & bin2Hex(s(i))))   'base64 alaphebet
    Next
        
    ReDim Preserve s(UBound(s) + pad)
    For i = 0 To pad - 1       'then remove them before processing?
       s(UBound(s) - i) = "="
    Next
    
    divs = Split(calcDivs(UBound(s), 72), ",") 'wrap characters at 72 chars
    For i = 1 To divs(0)                       'to conform to quoted
      s(i * 72) = s(i * 72) & vbCrLf           'printable standard (has to be < 76)
    Next                                       'but needs to have a multipul of 3 because of our chunked processing
    
    b64EncodeEngine = s 'BASE 1 string arrary of characters
Exit Function
warn: MsgBox "Err in B64EncodeEngine. This function accepts (and returns) only base 1 string arrays of individual characters" & vbCrLf & vbCrLf & Err.description
End Function

Private Function b64DecodeEngine(it)
 On Error GoTo warn
  If Not lib.Init Then initAlpha
  
  Dim s() As String        'it = BASE 1 string array of characters
  ReDim s(1 To UBound(it)) 'returns BASE 1 string array of characters
  Dim str As String
  pad = 0
  
  For i = 0 To 1                        'only last two vals could be pads
    If it(UBound(it) - i) = "=" Then      'if it is a pad then we will have to
      pad = pad + 1                     'remove it from the array and then
    End If                              'remove 2*pad bits from the binary
  Next                                  'stream latter(added 2bits/pad before)
  
  For i = 1 To UBound(it) - pad         'get base64 Ascii(val) of each char
    s(i) = Hex2Bin(Hex(b64Asc(it(i))))  'convert it to hex --> binary
    s(i) = Right(s(i), 6)               'only want last 6 chars, since 64 is
  Next                                  'max dec. value possible, are only trimming 0's
  
  str = Join(s, "")
  If pad Then str = Mid(str, 1, (Len(str) - pad * 2)) 'each pad effectivly adds 2 bits to stream
                                                      'now we get to remove them :)
  While Len(str) Mod 8 <> 0                           'rember was padded to encode properly
    str = str & "0"                                   'if it isnt encoded right we can usally salvage
  Wend                                                'and decodes with a max of 2 null chrs on end
  
  ReDim s(Len(str) Mod 8)    'clears contents redims to new size
  s = segment(str, 8)        'returns base 1 array of strings 8chr per
                             
  For i = 1 To UBound(s)
     s(i) = Chr(Int("&H" & bin2Hex(s(i))))
  Next
            
  b64DecodeEngine = s 'returns BASE 1 string array of characters
Exit Function
warn: MsgBox "Err in B64DecodeEngine. This function accepts(and returns) only base 1 string arrays of individual characters" & vbCrLf & vbCrLf & Err.description
End Function

Private Function segment(str As String, div As Integer)
    Dim t() As String       'returns BASE 1 STRING ARRAY of str
    ReDim t(1)              'broken up w/ div characters per element
                            'make sure it is even divisible before!!
    
    For i = 1 To Len(str) Step div
        t(UBound(t)) = Mid(str, i, div)
        If i < Len(str) - div Then         'you dont know how much debugging it
           ReDim Preserve t(UBound(t) + 1)  'took to find why I was always getting
        End If                              '1 extra null byte! ubound(t) was returning
    Next                                    '1 to many elements! last one was null :0`~_

    segment = t
End Function

Private Function calcDivs(maxsz As Long, division As Integer) As String
        sz = maxsz  'using maxsz directly changed its val in calling fx!!
        tmp = 0     'returns (max_whole_divisions,remainder)
        While sz >= division
           sz = sz - division
           tmp = tmp + 1
        Wend
        calcDivs = tmp & "," & sz
End Function

Private Sub ketchup(revs As Integer)
    For i = 0 To revs * 2
      DoEvents
    Next
End Sub

'************   base conversion *****************
Private Function bin2Hex(it As String) As String
  Dim t() As String    'it=binary val as string
  Short = 8 - Len(it)  'returns 2 chr hex string
  ReDim t(3)           'because segment is base1
  
  If Short Then         'need 8 char string to test
     For j = 1 To Short 'pad front with nulls
         it = "0" & it  '(doesnt change value)
     Next
  End If
  
  t = segment(it, 4) 'segment returns base 1 array
  For i = 1 To 2
    For j = 0 To 15
       If t(i) = lib.binHex(j) Then t(i) = lib.hexChr(j): Exit For
    Next
  Next

  bin2Hex = Join(t, "")
End Function

Private Function Hex2Bin(it As String) As String
  Dim tmp As String  'it = 2 char hex string
  If Len(it) = 1 Then it = "0" & it 'need 01 not 1 for val=1
   For i = 1 To 2
      ch = Mid(it, i, 1)
      If IsNumeric(ch) Then
        tmp = tmp & lib.binHex(ch)
      Else
        tmp = tmp & lib.binHex((Asc(ch) - 65 + 10))
      End If      'chr A--> asc65 -->hex chr 10
    Next
  Hex2Bin = tmp
End Function

Private Function b64Asc(it) As Integer
   Start = Asc(it)
   If Start > 64 And Start < 91 Then
      Start = 0
   ElseIf Start > 96 And Start < 123 Then
      Start = 26
   Else
      Start = 52
   End If
   
   For i = Start To 64
     If InStr(1, lib.b64Chr(i), it, vbBinaryCompare) > 0 Then
        b64Asc = i
        Exit For
     End If
   Next
End Function

'**************  end base conversion *************


Private Sub initAlpha()
  With lib
    'b64Alaphabet array
    For i = 0 To 25 '1-25 --> A-Z
     .b64Chr(i) = Chr(65 + i)
    Next
    For i = 26 To 51 '26-51 --> a-z
     .b64Chr(i) = Chr(97 + (i - 26))
    Next
    For i = 0 To 9   '52-61 --> 0-9
     .b64Chr(52 + i) = i
    Next
    .b64Chr(62) = "+"
    .b64Chr(63) = "/"
    .b64Chr(64) = "=" 'since orig val mod 3 must =0 these are the pads

    'hex-->binary array
    .binHex(0) = "0000":   .binHex(8) = "1000"
    .binHex(1) = "0001":   .binHex(9) = "1001"
    .binHex(2) = "0010":   .binHex(10) = "1010"
    .binHex(3) = "0011":   .binHex(11) = "1011"
    .binHex(4) = "0100":   .binHex(12) = "1100":
    .binHex(5) = "0101":   .binHex(13) = "1101"
    .binHex(6) = "0110":   .binHex(14) = "1110":
    .binHex(7) = "0111":   .binHex(15) = "1111"
    
    'hex alaphebet by index array
    .hexChr(0) = "0": .hexChr(6) = "6":  .hexChr(11) = "B"
    .hexChr(1) = "1": .hexChr(7) = "7":  .hexChr(12) = "C"
    .hexChr(2) = "2": .hexChr(8) = "8":  .hexChr(13) = "D"
    .hexChr(3) = "3": .hexChr(9) = "9":  .hexChr(14) = "E"
    .hexChr(4) = "4": .hexChr(10) = "A": .hexChr(15) = "F"
    .hexChr(5) = "5"
    
    .Init = True
  End With
End Sub









