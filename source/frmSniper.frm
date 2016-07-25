VERSION 5.00
Begin VB.Form frmSniper 
   Caption         =   "Drag & Drop Cross Hairs over IE window you wish to examine"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboColor 
      Height          =   315
      ItemData        =   "frmSniper.frx":0000
      Left            =   4140
      List            =   "frmSniper.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1380
      Width           =   1035
   End
   Begin VB.CommandButton cmdHighLightFindWord 
      Caption         =   "Color Find"
      Height          =   315
      Left            =   5220
      TabIndex        =   14
      ToolTipText     =   "color all instances of word in find box with selected color"
      Top             =   1380
      Width           =   1095
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      ToolTipText     =   "find this text in source"
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   1380
      Width           =   795
   End
   Begin VB.TextBox txtReplace 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "replace the find text with this"
      Top             =   1380
      Width           =   1035
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Colorize"
      Height          =   315
      Left            =   7260
      TabIndex        =   9
      ToolTipText     =   "Defaultsource highlighting scheme"
      Top             =   1380
      Width           =   795
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap"
      Height          =   255
      Left            =   6420
      TabIndex        =   8
      ToolTipText     =   "word wrap"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Parent Window"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   1875
   End
   Begin VB.TextBox txtCookie 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   7035
   End
   Begin VB.TextBox txtHref 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   540
      Width           =   7035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3900
      Top             =   60
   End
   Begin VB.TextBox txtHwnd 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin WebSleuth.RTF rtf 
      Height          =   2775
      Left            =   60
      TabIndex        =   6
      Top             =   1800
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   4895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cookie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "hWnd"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "href"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   390
   End
   Begin VB.Image imgSniper 
      Height          =   480
      Left            =   3360
      Picture         =   "frmSniper.frx":002C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmSniper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Functions related to the cursor position and its hwnd
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'User-defined type needed by the GetCursorPos function
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private ieDoc As IHTMLDocument

Private Sub cmdUpdate_Click()
    'ieDoc.body.innerHTML = rtf.text
    ieDoc.body.innerHTML = frmMain.wb.Document.body.innerHTML
End Sub

Private Sub chkWrap_Click()
    rtf.WordWrap = CBool(chkWrap.value)
End Sub

Private Sub cmdHighlight_Click()
    LockWindowUpdate rtf.hWnd
    rtf.highlightHtml
    LockWindowUpdate 0&
End Sub

Private Sub cmdHighLightFindWord_Click()
    LockWindowUpdate rtf.hWnd
    c = Array(vbRed, vbBlue, &HC000C0, &H808000)
    rtf.SetColor txtFind, CLng(c(cboColor.ListIndex)), , True
    rtf.ScrollToTop
    LockWindowUpdate 0&
End Sub

Private Sub cmdFind_Click()
    If txtFind <> rtf.FindString Then
        rtf.FindString = txtFind
        rtf.find
    Else
        rtf.findNext
    End If
End Sub

Private Sub cmdReplace_Click()
    rtf.ReplaceText txtFind, txtReplace
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtf.Height = Me.Height - rtf.Top - 150
    rtf.Width = Me.Width - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ieDoc = Nothing
End Sub

Private Sub Timer1_Timer()
    Dim p As POINTAPI
    GetCursorPos p
    txtHwnd = WindowFromPoint(p.X, p.Y)
End Sub

Private Sub imgSniper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 99 'custom
    Screen.MouseIcon = LoadResPicture("sniper.ico", vbResIcon)
    Timer1.Enabled = True
End Sub

Private Sub imgSniper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
    Timer1.Enabled = False
    If modSniper.IsIEServerWindow(CLng(txtHwnd)) Then
        Set ieDoc = modSniper.IEDOMFromhWnd(CLng(txtHwnd))
        frmMain.TransferObject ieDoc
        'rtf.text = ieDoc.body.innerHTML
        'txtHref = ieDoc.location
        'txtCookie = ieDoc.cookie
    Else
        txtHwnd = "Not Valid IE Window"
    End If
End Sub

Private Sub txtHref_GotFocus()
     txtHref.SelStart = 0
     txtHref.SelLength = Len(txtHref.text)
End Sub

