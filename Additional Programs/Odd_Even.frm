VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Even or Odd"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EVEN/ODD"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function iseven(N As Integer) As Boolean
    If N Mod 2 = 0 Then
        iseven = True
    Else
        iseven = False
    End If
End Function


Private Sub Command1_Click()
    If iseven(Text1.Text) = True Then
        MsgBox ("The number you typed is even")
    Else
        MsgBox ("the number you typed is odd")
    End If
End Sub

