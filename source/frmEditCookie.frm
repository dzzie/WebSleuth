VERSION 5.00
Begin VB.Form frmEditCookie 
   Caption         =   "Windows Edit Cookie Api  - Play with get set can try any url for cookie on machine"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "form1"
   ScaleHeight     =   3705
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Add Exp"
      Height          =   330
      Left            =   6300
      TabIndex        =   7
      Top             =   945
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SET"
      Height          =   330
      Left            =   6300
      TabIndex        =   4
      Top             =   525
      Width           =   1170
   End
   Begin VB.TextBox txtCookieValue 
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
      Height          =   2010
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1575
      Width           =   7470
   End
   Begin VB.TextBox txtCookieName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   2
      Top             =   525
      Width           =   4635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET"
      Height          =   330
      Left            =   6300
      TabIndex        =   1
      Top             =   105
      Width           =   1170
   End
   Begin VB.TextBox txtUrl 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   0
      Text            =   "http://yahoo.com"
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label Label3 
      Caption         =   "for local proxy soon"
      Height          =   225
      Left            =   6195
      TabIndex        =   9
      Top             =   1365
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   $"frmEditCookie.frx":0000
      Height          =   615
      Left            =   105
      TabIndex        =   8
      Top             =   945
      Width           =   6045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CookieName"
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
      Left            =   105
      TabIndex        =   6
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Page URL"
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
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   915
   End
End
Attribute VB_Name = "frmEditCookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'yay msdn! betcha never knew these existed :D
 
Private Declare Function InternetGetCookieA Lib "wininet.dll" _
  (ByVal lpszUrlName As String, _
     ByVal lpszCookieName As String, _
     ByVal lpszCookieData As String, _
     lpdwSize As Long _
  ) As Long

'BOOL InternetGetCookie(
'    LPCTSTR lpszUrlName,
'    LPCTSTR lpszCookieName,
'    LPTSTR lpszCookieData,
'    LPDWORD lpdwSize
');

Private Declare Function InternetSetCookieA Lib "wininet.dll" _
  (ByVal lpszUrl As String, _
   ByVal lpszCookieName As String, _
   ByVal lpszCookieData As String _
  ) As Long
  
'BOOL InternetSetCookie(
'    LPCTSTR lpszUrl,
'    LPCTSTR lpszCookieName,
'    LPCTSTR lpszCookieData
');

'  // Create a session cookie.
'bReturn = InternetSetCookie("http://www.adventure_works.com", NULL,
'            "TestData = Test");
'
'// Create a persistent cookie.
'bReturn = InternetSetCookie("http://www.adventure_works.com", NULL,
'             "TestData = Test; expires = Sat, 01-Jan-2000 00:00:00 GMT");
             
             

  
Private Sub Command1_Click() 'cookie get
     Dim Ret As Long
     Dim buffer As Long
     Dim contents As String
      
      'if buffer=0 then it will set buffer to size needed
      Ret = InternetGetCookieA(txtUrl, txtCookieName, contents, buffer)
      If Ret = 0 Or buffer < 1 Then Exit Sub
    
      contents = String(buffer, Chr(0))
      buffer = buffer - 1
      
      InternetGetCookieA txtUrl, txtCookieName, contents, buffer
      txtCookieValue = Left$(contents, InStrB(1, contents, Chr(0), vbBinaryCompare))
      
End Sub

Private Sub Command2_Click()
    InternetSetCookieA txtUrl, txtCookieName, txtCookieValue
    txtCookieValue = Empty
End Sub

Sub LoadFormFromUrl(url)
    txtUrl = url
    Command1_Click
    Me.Show
End Sub

Private Sub Command3_Click()
  If txtCookieValue <> Empty Then
    txtCookieValue = txtCookieValue & "; expires = Sat, 01-Jan-2003 00:00:00 GMT"
  End If
End Sub
