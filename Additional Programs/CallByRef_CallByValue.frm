VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Call By Value / Calll By Reference"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton callbyref 
      Caption         =   "Call By Reference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton callbyval 
      Caption         =   "Call By Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox secnum 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox firstnum 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblsecnum 
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblfirstnum 
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Second Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "First Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub callbyref_Click()
Dim a1 As Integer, b1 As Integer
a1 = firstnum.Text
b1 = secnum.Text
Call swapbyref(a1, b1)
lblfirstnum.Caption = a1
lblsecnum.Caption = b1
lblmsg.Caption = "swapped by reference"
End Sub


Private Sub callbyval_Click()
Dim a1, b1 As Integer
a1 = firstnum.Text
b1 = secnum.Text
Call swapbyval(a1, b1)
lblfirstnum.Caption = a1
lblsecnum.Caption = b1
lblmsg.Caption = "swapped by value"
End Sub

Public Sub swapbyval(ByVal a As Integer, ByVal b As Integer)
Dim temp As Integer
temp = a
a = b
b = temp
End Sub

Public Sub swapbyref(ByRef a As Integer, ByRef b As Integer)
Dim temp As Integer
temp = a
a = b
b = temp
End Sub


Private Sub exit_Click()
End
End Sub
