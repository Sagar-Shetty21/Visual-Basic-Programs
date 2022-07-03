VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple Calculator"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton equal 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   17
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton digit 
      Caption         =   "0"
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
      Index           =   3
      Left            =   1440
      TabIndex        =   16
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton clear 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton div 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton dot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton star 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "3"
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
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "2"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "1"
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
      Index           =   6
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "6"
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
      Index           =   7
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "5"
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
      Index           =   8
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "4"
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
      Index           =   9
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "9"
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
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "8"
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
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton digit 
      Caption         =   "7"
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
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label display 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operand As Double, operand2 As Double
Dim op1 As Double, op2 As Double
Dim operator As String
Dim cleardisplay As Boolean


Private Sub clear_Click()
    display.Caption = ""
End Sub



Private Sub div_Click()
    op1 = Val(display.Caption)
    operator = "/"
    display.Caption = ""
End Sub


Private Sub dot_Click()
    If InStr(display.Caption, ".") Then
        Exit Sub
    Else
        display.Caption = display.Caption + "."
    End If
End Sub



Private Sub equal_Click()
    Dim result As Double
    op2 = Val(display.Caption)
    If operator = "+" Then
        result = add(ByVal op1, ByVal op2)
    End If
    If operator = "*" Then
        result = mul(ByVal op1, ByVal op2)
    End If
    If operator = "-" Then
        result = subtract(ByVal op1, ByVal op2)
    End If
    If operator = "/" Then
        result = divi(ByVal op1, ByVal op2)
    End If
    display.Caption = result
End Sub


Private Sub star_Click()
    op1 = Val(display.Caption)
    operator = "*"
    display.Caption = ""
End Sub

Private Sub minus_Click()
    op1 = Val(display.Caption)
    operator = "-"
    display.Caption = ""
End Sub

Private Sub plus_Click()
    op1 = Val(display.Caption)
    operator = "+"
    display.Caption = ""
End Sub

Private Sub digit_Click(index As Integer)
    If cleardisplay Then
        display.Caption = ""
        cleardisplay = False
    End If
    display.Caption = display.Caption + digit(index).Caption
End Sub

Private Function add(ByVal operand1 As Double, ByVal operand2 As Double) As Double
    add = operand1 + operand2
End Function

Private Function subtract(ByVal operand1 As Double, ByVal operand2 As Double) As Double
    subtract = operand1 - operand2
End Function

Private Function mul(ByVal operand1 As Double, ByVal operand2 As Double) As Double
    mul = operand1 * operand2
End Function

Private Function divi(ByVal operand1 As Double, ByVal operand2 As Double) As Double
    divi = operand1 / operand2
End Function

