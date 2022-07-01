VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Average Form"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox textavg 
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Text            =   "0"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txttotal 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtsecond 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtfirst 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Average"
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
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Total"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Second Number"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "First Number"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    txttotal = Val(txtfirst) + (txtsecond)
End Sub

Private Sub Command2_Click()
    textavg = txttotal / 2
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
