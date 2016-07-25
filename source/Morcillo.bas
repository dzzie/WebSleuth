Attribute VB_Name = "modSniper"
Option Explicit

'this amazing module from:
'
'http://www.domaindlx.com/e_morcillo/default.asp
'e_morcillo@yahoo.com
'
'excellent work and so appreciated that it is open source :)
'please everyone check his site out...it is simply amazing !

'All controls and TypeLibs in this site are freely distributable.
'You are free to use the source code in your applications, but it's
'forbidden to distribute the code in sample projects, web sites or
'news posts without mention the author's name (my name :-) ) or
'this site URL.

'If you have any comment, bug report or question feel free to mail
'it to Eduardo A. Morcillo

'
' Requires: reference to "Microsoft HTML Object Library"
'

Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Public Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" ( _
      ByVal hWnd As Long, _
      ByVal lpClassName As String, _
      ByVal nMaxCount As Long) As Long

Public Declare Function EnumChildWindows Lib "user32" ( _
      ByVal hWndParent As Long, _
      ByVal lpEnumFunc As Long, _
      lParam As Long) As Long

Public Declare Function RegisterWindowMessage Lib "user32" _
   Alias "RegisterWindowMessageA" ( _
      ByVal lpString As String) As Long

Public Declare Function SendMessageTimeout Lib "user32" _
   Alias "SendMessageTimeoutA" ( _
      ByVal hWnd As Long, _
      ByVal msg As Long, _
      ByVal wParam As Long, _
      lParam As Any, _
      ByVal fuFlags As Long, _
      ByVal uTimeout As Long, _
      lpdwResult As Long) As Long
      
Public Const SMTO_ABORTIFHUNG = &H2

Public Declare Function ObjectFromLresult Lib "oleacc" ( _
      ByVal lResult As Long, _
      riid As UUID, _
      ByVal wParam As Long, _
      ppvObject As Any) As Long

Public Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" ( _
      ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long

'
' IEDOMFromhWnd
'
' Returns the IHTMLDocument interface from a WebBrowser window
'
' hWnd - Window handle of the control
'
Function IEDOMFromhWnd(ByVal hWnd As Long) As HTMLDocument
Dim IID_IHTMLDocument As UUID
Dim hWndChild As Long
Dim spDoc As IUnknown
Dim lRes As Long
Dim lMsg As Long
Dim hr As Long

   If hWnd <> 0 Then
      
      If Not IsIEServerWindow(hWnd) Then
      
         ' Get 1st child IE server window
         EnumChildWindows hWnd, AddressOf EnumChildProc, hWnd
         
      End If
      
      If hWnd <> 0 Then
            
            ' Register the message
            lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
            
            ' Get the object
            Call SendMessageTimeout(hWnd, lMsg, 0, 0, _
                 SMTO_ABORTIFHUNG, 1000, lRes)

            If lRes Then
               
               ' Initialize the interface ID
               With IID_IHTMLDocument
                  .Data1 = &H626FC520
                  .Data2 = &HA41E
                  .Data3 = &H11CF
                  .Data4(0) = &HA7
                  .Data4(1) = &H31
                  .Data4(2) = &H0
                  .Data4(3) = &HA0
                  .Data4(4) = &HC9
                  .Data4(5) = &H8
                  .Data4(6) = &H26
                  .Data4(7) = &H37
               End With
               
               ' Get the object from lRes
               hr = ObjectFromLresult(lRes, IID_IHTMLDocument, _
                     0, IEDOMFromhWnd)
               
            End If

      End If
      
   End If

End Function

Function EnumChildProc(ByVal hWnd As Long, lParam As Long) As Long
   
   If IsIEServerWindow(hWnd) Then
      lParam = hWnd
   Else
      EnumChildProc = 1
   End If
   
End Function

Function IsIEServerWindow(ByVal hWnd As Long) As Boolean
Dim lRes As Long
Dim sClassName As String

   ' Initialize the buffer
   sClassName = String$(100, 0)
   
   ' Get the window class name
   lRes = GetClassName(hWnd, sClassName, Len(sClassName))
   sClassName = Left$(sClassName, lRes)
   
   IsIEServerWindow = StrComp(sClassName, _
                      "Internet Explorer_Server", _
                      vbTextCompare) = 0
   
End Function



