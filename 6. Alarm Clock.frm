VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alarm Clock Application"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "SET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
