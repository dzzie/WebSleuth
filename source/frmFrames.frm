VERSION 5.00
Begin VB.Form frmFrames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Frames Layout      NOTE:  Empty Index = parent frame"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Frame Index String"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1875
   End
End
Attribute VB_Name = "frmFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret() As String 'parentName/subname/subname2
Dim idx() As String 'parentIndex,subframeIndex,subframe2Index
Dim waiting As Boolean

Sub ListFrames()
    Erase ret()
    Erase idx()
    List1.Clear
    GenList frmMain.wb.Document, Empty, Empty
    If Not aryIsEmpty(ret) Then
        For i = 0 To UBound(ret)
            List1.AddItem ret(i)
        Next
    End If
    If List1.ListCount > 0 Then Me.Show
End Sub

Private Sub GenList(d As HTMLDocument, parentStr, indexStr)
        On Error Resume Next
        For i = 0 To d.frames.length - 1
             push ret(), parentStr & "/" & d.frames(i).Name
             push idx(), indexStr & "," & i
             If d.frames(i).Document.frames.length > 0 Then
                GenList d.frames(i).Document, _
                        parentStr & "/" & d.frames(i).Name, _
                        indexStr & "," & i
             End If
        Next
End Sub

Function AnyAccessibleFrames() As Boolean
    Call ListFrames
    If List1.ListCount > 0 Then AnyAccessibleFrames = True
End Function

Function ReturnFrameIndex()
    Call ListFrames
    waiting = True
    
    If List1.ListCount = 0 Then waiting = False
    
    While waiting
        DoEvents
        Sleep 100
    Wend
    
    ReturnFrameIndex = Text1
    
    Me.Hide
    Text1 = Empty
End Function

Function ReturnFrameObject() As HTMLDocument
    
        Dim f As HTMLDocument
        
        strindex = frmFrames.ReturnFrameIndex
        If strindex <> Empty Then
            ret = Split(strindex, ",")
            Select Case UBound(ret)
                Case 0: MsgBox "no frames?!"
                Case 1: Set f = frmMain.wb.Document.frames(ret(1)).Document
                Case 2: Set f = frmMain.wb.Document.frames(ret(1)).Document.frames(ret(2)).Document
                Case 3: Set f = frmMain.wb.Document.frames(ret(1)).Document.frames(ret(2)).Document.frames(ret(3)).Document
                Case Else: MsgBox "OK lets not be ridiclous i dont have an eval here :P, frames nested deeper than 3 layers which i didnt account for sorry :("
            End Select
        End If
        
        Set ReturnFrameObject = f
        
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    waiting = False
End Sub

Private Sub List1_Click()
    On Error Resume Next
    Text1 = idx(List1.ListIndex)
End Sub

Private Sub Command1_Click()
    If waiting Then
        waiting = False
    Else
        Unload Me
    End If
End Sub

Private Sub List1_DblClick()
    List1_Click
    Command1_Click
End Sub
