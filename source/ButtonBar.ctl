VERSION 5.00
Begin VB.UserControl ButtonBar 
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ScaleHeight     =   2595
   ScaleWidth      =   3360
   Begin VB.PictureBox Img 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   4
      Left            =   3180
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox Img 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   3
      Left            =   2340
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox Img 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   2
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox Img 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   1
      Left            =   780
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox Img 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   0
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3240
      Top             =   300
   End
   Begin VB.Image cImg 
      Height          =   435
      Index           =   4
      Left            =   2040
      Picture         =   "ButtonBar.ctx":0000
      Top             =   1980
      Width           =   840
   End
   Begin VB.Image cImg 
      Height          =   435
      Index           =   3
      Left            =   2880
      Picture         =   "ButtonBar.ctx":10A4
      Top             =   1980
      Width           =   450
   End
   Begin VB.Image cImg 
      Height          =   435
      Index           =   2
      Left            =   1560
      Picture         =   "ButtonBar.ctx":227E
      Top             =   1980
      Width           =   450
   End
   Begin VB.Image cImg 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "ButtonBar.ctx":3869
      Top             =   1980
      Width           =   675
   End
   Begin VB.Image cImg 
      Height          =   435
      Index           =   0
      Left            =   720
      Picture         =   "ButtonBar.ctx":4892
      Top             =   1980
      Width           =   840
   End
   Begin VB.Image dImg 
      Height          =   435
      Index           =   4
      Left            =   2040
      Picture         =   "ButtonBar.ctx":58B8
      Top             =   780
      Width           =   840
   End
   Begin VB.Image dImg 
      Height          =   435
      Index           =   3
      Left            =   2880
      Picture         =   "ButtonBar.ctx":6669
      Top             =   780
      Width           =   450
   End
   Begin VB.Image dImg 
      Height          =   435
      Index           =   2
      Left            =   1560
      Picture         =   "ButtonBar.ctx":7470
      Top             =   780
      Width           =   450
   End
   Begin VB.Image dImg 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "ButtonBar.ctx":8365
      Top             =   780
      Width           =   675
   End
   Begin VB.Image dImg 
      Height          =   435
      Index           =   0
      Left            =   720
      Picture         =   "ButtonBar.ctx":8FB6
      Top             =   780
      Width           =   840
   End
   Begin VB.Image eImg 
      Height          =   435
      Index           =   4
      Left            =   1980
      Picture         =   "ButtonBar.ctx":9CEF
      ToolTipText     =   "Forward"
      Top             =   1380
      Width           =   840
   End
   Begin VB.Image eImg 
      Height          =   435
      Index           =   3
      Left            =   2820
      Picture         =   "ButtonBar.ctx":AD88
      ToolTipText     =   "Refresh"
      Top             =   1380
      Width           =   450
   End
   Begin VB.Image eImg 
      Height          =   435
      Index           =   2
      Left            =   1500
      Picture         =   "ButtonBar.ctx":C230
      ToolTipText     =   "Stop"
      Top             =   1380
      Width           =   450
   End
   Begin VB.Image eImg 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "ButtonBar.ctx":D9F6
      ToolTipText     =   "Go !"
      Top             =   1380
      Width           =   675
   End
   Begin VB.Image eImg 
      Height          =   435
      Index           =   0
      Left            =   660
      Picture         =   "ButtonBar.ctx":EA13
      ToolTipText     =   "Navigate Back"
      Top             =   1380
      Width           =   840
   End
End
Attribute VB_Name = "ButtonBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'mouseover animations adds overhead, not a whole lot
'but still it annoys me to think abotu it :P
'on a pentium 2 235mhz the mouseover monitoring and timer
'only increased processor load 2-4% which isnt bad but still.


Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Enum RunTimeConstants
    NoMouseOver = 0
    MouseOver = 1
End Enum

Private UseMouseOvers As Boolean

Event Click(index As Integer)

Property Let RunTimeStyle(ByVal choice As RunTimeConstants)
    UseMouseOvers = CBool(choice)
    UserControl_Initialize
End Property

Property Get RunTimeStyle() As RunTimeConstants
    RunTimeStyle = IIf(UseMouseOvers, True, False)
End Property

Private Sub Img_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img(index).Picture = cImg(index).Picture
    Timer1.Enabled = False
End Sub

Private Sub Img_MouseUp(i As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim pt As POINTAPI
    GetCursorPos pt

    If Img(i).hWnd <> WindowFromPoint(pt.x, pt.Y) Then
        Img(i).Picture = dImg(i).Picture
    Else
        Img(i).Picture = dImg(i).Picture
        RaiseEvent Click(i)
    End If

    If UseMouseOvers Then Timer1.Enabled = True
       
End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI
    GetCursorPos pt

    For i = 0 To Img.UBound
        If Img(i).hWnd <> WindowFromPoint(pt.x, pt.Y) Then
            Img(i).Picture = dImg(i).Picture
        Else
            Img(i).Picture = eImg(i).Picture
        End If
    Next
End Sub


Private Sub UserControl_Initialize()
    
    If IsIde Then UseMouseOvers = False
    
    If UseMouseOvers Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
    
    For i = 0 To Img.UBound
         'load default disabled images and put buttons in right place
         Img(i).Picture = dImg(i).Picture
         Img(i).Move eImg(i).Left
         Img(i).ToolTipText = eImg(i).ToolTipText
    Next
        
End Sub

Private Function IsIde() As Boolean
On Error GoTo poo
    Debug.Print (1 / 0)
    IsIde = False
Exit Function
poo: IsIde = True
End Function
