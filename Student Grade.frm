VERSION 5.00
Begin VB.Form Student_Grade 
   Caption         =   "Student Grade"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton grade 
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox marks 
      Alignment       =   2  'Center
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
      Left            =   5040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label ans 
      Alignment       =   2  'Center
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
      Left            =   5040
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Grade"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter the average marks of the student"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label header 
      Caption         =   "FIND THE GRADE OF A STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Student_Grade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear_Click()
    marks.Text = ""
    ans.Caption = ""
    marks.SetFocus
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub grade_Click()
    If marks > 75 Then
        ans.Caption = "A"
    ElseIf marks > 65 Then
        ans.Caption = "B"
    ElseIf marks > 55 Then
        ans.Caption = "C"
    ElseIf marks >= 45 Then
        ans.Caption = "S"
    Else
        ans.Caption = "FAIL"
    End If
End Sub
