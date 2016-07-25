VERSION 5.00
Begin VB.Form frmLogOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log options"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "Meta Tags"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   1560
      TabIndex        =   5
      Top             =   900
      Width           =   1515
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "innerHtml"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   1395
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "Scripts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   1215
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "Images"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CheckBox ChkLogOptions 
      Caption         =   "Cookies"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmLogOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
