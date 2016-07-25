VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPUT 
   Caption         =   "HTTP PUT & DELETE"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   390
      Left            =   2430
      TabIndex        =   13
      Top             =   1515
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "PUT.frx":0000
      Left            =   885
      List            =   "PUT.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1500
      Width           =   1350
   End
   Begin VB.TextBox t1 
      Height          =   300
      Index           =   4
      Left            =   825
      TabIndex        =   10
      Text            =   "www.bad-things.com"
      Top             =   1110
      Width           =   2685
   End
   Begin VB.TextBox t1 
      Height          =   300
      Index           =   3
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   750
      Width           =   2685
   End
   Begin VB.TextBox t1 
      Height          =   300
      Index           =   2
      Left            =   825
      TabIndex        =   7
      Top             =   405
      Width           =   2685
   End
   Begin VB.TextBox t1 
      Height          =   330
      Index           =   1
      Left            =   2955
      TabIndex        =   4
      Text            =   "80"
      Top             =   45
      Width           =   540
   End
   Begin VB.TextBox t1 
      Height          =   300
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   60
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8610
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox t2 
      Height          =   1755
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2205
      Width           =   3450
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      Caption         =   "Will Send Raw Text Not Generated Headers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   3990
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Label lblLock 
      AutoSize        =   -1  'True
      Caption         =   "UnLock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2940
      TabIndex        =   14
      Top             =   1995
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Method"
      Height          =   195
      Index           =   2
      Left            =   195
      TabIndex        =   12
      Top             =   1575
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Write Dir"
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   8
      Top             =   465
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP : Port"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   570
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      Caption         =   "Raw Headers"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host Val:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   1185
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Path"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   825
      Width           =   615
   End
End
Attribute VB_Name = "frmPUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim incoming As String
Dim header As String
Dim filePath As String

Dim body As Variant

Const fileMarker = "<File Content>"



Private Sub Combo1_Click()
 If t1(3) <> Empty Then GenerateHeaders
End Sub

Private Sub Command1_Click()
  Clipboard.Clear
  Clipboard.SetText t2
  incoming = ""
  If Winsock1.State <> sckClosed Then Winsock1.Close
  Winsock1.Connect t1(0), t1(1)
  Me.Caption = "Connecting..."
End Sub

Private Sub Form_Load()
  Combo1.ListIndex = 0
End Sub

Private Sub lblLock_Click()
  With lblLock
        If .Caption = "UnLock" Then
            .Caption = "Lock"
            t2.Locked = False
            lblWarn.Visible = True
        Else
            .Caption = "UnLock"
            t2.Locked = True
            t2 = header
            lblWarn.Visible = False
        End If
  End With
End Sub

Private Sub t1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 3 Then
        MsgBox "Just drop the file you want to upload in here dont try to type in teh path...", vbInformation
        t1(3) = Empty
   End If
End Sub

Private Sub t1_LostFocus(Index As Integer)
  If Index = 3 Then GenerateHeaders
End Sub

Private Sub t1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   filePath = Data.Files(1)
   t1(Index) = FileNameFromPath(filePath)
   GenerateHeaders
   lblStat = "Raw Headers"
End Sub

Private Sub t2_DblClick()
  t2 = ""
  If lblWarn.Visible = True Then t2 = Clipboard.GetText
End Sub

Private Sub Winsock1_Close()
  On Error Resume Next
  Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
 If lblWarn.Visible = True Then
    Dim sendBody As Boolean
    If InStr(t2, fileMarker) > 0 Then
        t2 = Replace(t2, fileMarker, Empty)
        sendBody = True
    End If
    Winsock1.SendData t2 & IIf(sendBody = True, body, Empty) & vbCrLf
 Else
    Winsock1.SendData header & body & vbCrLf
 End If
 Me.Caption = "Sending..."
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim tmp As String
  Winsock1.GetData tmp, vbString, bytesTotal
  incoming = incoming & tmp
  lblStat = "Server Response"
  Me.Caption = "Ready- No Connection"
  t2 = incoming
End Sub

Sub GenerateHeaders()
  On Error Resume Next
  header = ""
  body = ""
  
  If Right(t1(2), 1) = "/" Then t1(2) = Mid(t1(2), 1, Len(t1(2)) - 1)
  If t1(2) <> Empty And Left(t1(2), 1) <> "/" Then t1(2) = "/" & t1(2)
  
  Select Case Combo1.ListIndex
     Case 0 ' PUT
        tmp = "PUT " & t1(2) & "/" & t1(3) & " HTTP/1.1" & vbCrLf
        tmp = tmp & "Content-Length: " & FileLen(filePath) & vbCrLf
        tmp = tmp & "Host: " & t1(4) & vbCrLf & vbCrLf
        
        header = tmp
        body = getFile(filePath)
        t2 = tmp & fileMarker
        
     Case 1 'DELETE
        tmp = "DELETE " & t1(2) & "/" & t1(3) & " HTTP/1.1" & vbCrLf
        tmp = tmp & "Host:" & t1(4) & vbCrLf & vbCrLf
        header = tmp
        t2 = tmp
        
  End Select
End Sub

Private Function getFile(filename)
  f = FreeFile
  Temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   getFile = Temp
End Function

Public Function FileNameFromPath(fullPath) As String
    If InStr(fullPath, "\") > 0 Then
        tmp = Split(fullPath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox Description
End Sub
