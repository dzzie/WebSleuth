VERSION 5.00
Begin VB.Form frmProxy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Proxy Manager"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2835
      TabIndex        =   7
      Top             =   945
      Width           =   855
   End
   Begin WebSleuth.reg reg 
      Left            =   0
      Top             =   0
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   525
      Width           =   3375
   End
   Begin VB.CommandButton cmdManage 
      Caption         =   "Manage List"
      Height          =   330
      Left            =   1365
      TabIndex        =   4
      Top             =   945
      Width           =   1170
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Done"
      Height          =   330
      Left            =   3675
      TabIndex        =   3
      Top             =   945
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Proxy"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Proxy"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "Current Proxy"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim pth As String, pList As String, proxies() As String

Private Sub Form_Load()
    pth = "Software\Microsoft\windows\CurrentVersion\Internet Settings"
    pList = App.path & "\ProxyList.txt"  'below for use in IDE
    If Not FileExists(pList) Then pList = App.path & ".\..\ProxyList.txt"
    If FileExists(pList) Then
        proxies() = Split(ReadFile(pList), vbCrLf)
        If Not aryIsEmpty(proxies) Then
            For i = 0 To UBound(proxies)
                If proxies(i) <> Empty Then Combo1.AddItem proxies(i)
            Next
        End If
        Combo1.ListIndex = 0
    Else
        MsgBox "Oops..couldnt find default proxy list" & vbCrLf & vbCrLf & pList
    End If
    Text1 = reg.ReadValue(HKEY_CURRENT_USER, pth, "ProxyServer")
    Check1 = reg.ReadValue(HKEY_CURRENT_USER, pth, "ProxyEnable")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdManage_Click()
    ShellnWait "notepad """ & pList, vbNormalFocus
    Combo1.Clear
    Form_Load
End Sub

Private Sub cmdOk_Click()
  On Error GoTo oops
    reg.SetValue HKEY_CURRENT_USER, pth, "ProxyServer", Text1, REG_SZ
    reg.SetValue HKEY_CURRENT_USER, pth, "ProxyEnable", Check1.value, REG_DWORD
        
    'this forces IE to recgonize proxy change
    Dim d As WebBrowser_V1
    Set d = New WebBrowser_V1
    Set d = Nothing
    Unload Me
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    Text1 = Combo1.List(Combo1.ListIndex)
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

