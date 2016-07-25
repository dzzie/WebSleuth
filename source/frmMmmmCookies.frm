VERSION 5.00
Begin VB.Form frmMmmmCookies 
   Caption         =   "Yummie Cookies ! "
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin WebSleuth.List List1 
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Right Click to Copy, Double Click to Analyze"
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "frmMmmmCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowCookie(cookie)
    List1.LoadArray BreakDownCookie(cookie)
End Sub

Private Sub Form_Load()
    n = "frmUmmie"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)

    SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
        Me.Top / 15, Me.Width / 15, _
        Me.Height / 15, Empty
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Height = Me.Height - 100
    List1.Width = Me.Width - 100
    'cause list height goes in steps
    Me.Height = List1.Height + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.ChkCookieOnTop.value = 0
    n = "frmUmmie"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
    End If
End Sub

Private Sub List1_DoubleClick()
    frmAnalyze.AnlyzeCookie frmMain.wb.Document.cookie
End Sub

Private Sub List1_RightClick()
    Screen.MousePointer = vbHourglass
    Me.caption = "Copying Cookie..."
    Sleep 200
    Clipboard.Clear
    Clipboard.SetText frmMain.wb.LocationURL & vbCrLf & vbCrLf & List1.GetListContents
    Sleep 200
    Me.caption = "Copy Complete..."
    Screen.MousePointer = vbDefault
End Sub
