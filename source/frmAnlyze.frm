VERSION 5.00
Begin VB.Form frmAnalyze 
   Caption         =   "Analyze Data"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   ScaleHeight     =   2865
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   2310
      Width           =   5580
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3150
         TabIndex        =   3
         Top             =   0
         Width           =   1170
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         TabIndex        =   2
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.TextBox Text1 
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
      Height          =   2115
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAnlyze.frx":0000
      Top             =   105
      Width           =   5580
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuBatchReplace 
         Caption         =   "Batch Replace"
      End
      Begin VB.Menu mnuEscape 
         Caption         =   "URL Decode"
         Index           =   0
      End
      Begin VB.Menu mnuEscape 
         Caption         =   "URL Encode"
         Index           =   1
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Base64 Decode Sel"
         Index           =   0
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Base64 Encode Sel"
         Index           =   1
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlugins 
         Caption         =   "Plugins"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private readyToReturn As Boolean
Private Cancel As Boolean


Private Sub mnuBatchReplace_Click()
    On Error GoTo oops
    it = InputBox("Enter Batch Replace String. Command format:" & vbCrLf & vbCrLf & "this->that,andThis->toThat" & vbCrLf & vbCrLf & "Note that CR is a keyword to insert a carriage return")
    If it = Empty Then Exit Sub
    Text1 = BatchReplace(Text1, Replace(it, "CR", vbCrLf))
    Exit Sub
oops: MsgBox Err.Description, vbInformation
End Sub

Private Sub mnuPlugins_Click(index As Integer)
     FirePluginEvent index, "frmAnalyze.mnuPlugins"
End Sub

Private Sub Command1_Click()
  readyToReturn = True
  Me.Hide
End Sub
Private Sub Command2_Click()
   Cancel = True
   Me.Hide
End Sub

Function AnlyzeCookie(it, Optional AndWait = True, Optional caption = "Anlyze Cookie: Note CRLF will be removed before return")
    readyToReturn = False
    Cancel = False
    
    it = Replace(it, "&", vbCrLf & "&")
    Text1 = Replace(it, ";", ";" & vbCrLf)
    Me.caption = caption
    Me.Show
    
    If AndWait Then
        While Not readyToReturn
            DoEvents: Sleep 100
            If Cancel Then Text1 = -1: readyToReturn = True
        Wend
    End If
    
    AnlyzeCookie = Replace(Text1, vbCrLf, Empty)
End Function


Function AnlyzeUrl(ByVal it, Optional AndWait = True, Optional showOnlyCgis = True, Optional caption = "Anlyze Link: Note CRLF will be removed before return")
    
    If showOnlyCgis Then
        If InStr(it, "?") < 1 And InStr(it, "&") < 1 Then: AnlyzeUrlAndWait = it: Exit Function
    End If
    
    readyToReturn = False
    Cancel = False
    Me.caption = caption
    Me.Show
    
    it = Replace(it, "?", vbCrLf & vbCrLf & "?")
    Text1 = Replace(it, "&", vbCrLf & "&")
    
    If AndWait Then
        While Not readyToReturn
            DoEvents: Sleep 100
            If Cancel = True Then Text1 = -1: readyToReturn = True
        Wend
    End If
    
    AnlyzeUrl = Replace(Text1, vbCrLf, Empty)
End Function

Sub ShowScript(it)
    Text1 = UnixToDos(it)
    Me.Show
End Sub

Private Sub mnuDecode_Click(index As Integer)
    If Text1.SelText = Empty Then Exit Sub
    If index = 1 Then Text1.SelText = b64Encode(Text1.SelText) _
    Else: Text1.SelText = b64Decode(Text1.SelText)
End Sub

Private Sub mnuEscape_Click(index As Integer)
    If Text1.SelText = Empty Then Exit Sub
    If index = 1 Then Text1.SelText = escape(Text1.SelText) _
    Else Text1.SelText = UnEscape(Text1.SelText)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Height = Me.Height - 1150
    Text1.Width = Me.Width - 300
    Frame.Top = Text1.Height + Text1.Top + 150
    Frame.Left = Me.Width - Frame.Width - 200
End Sub

Private Sub Form_Load()
    n = "frmanlyze"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'if the frmUnloads completly when they close it with the x then
    'any additions to the rt click menu from plugins will disappear!
    Cancel = 1
    Me.Hide
    
    n = "frmanlyze"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
    End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then ShowRtClkMenu Me, Text1, mnuPopup
End Sub
