VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Student Details"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cal 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4560
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
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
      Left            =   6840
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
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
      Left            =   6840
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Line Line10 
      X1              =   1320
      X2              =   6600
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line9 
      X1              =   1320
      X2              =   6600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line8 
      X1              =   1320
      X2              =   6600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line7 
      X1              =   1320
      X2              =   6600
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   6600
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line4 
      X1              =   6600
      X2              =   6600
      Y1              =   840
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   6600
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   840
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   6600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label123 
      Caption         =   "DIVISION"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label18 
      Caption         =   "PERCENTAGE"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label167 
      Caption         =   "%"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label11 
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
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label10 
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
      Left            =   4560
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "TOTAL MARKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "SUBJECT 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "SUBJECT 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "SUBJECT 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STUDENT DETAIL"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cal_Click()
    Dim a As Integer
    If (Text1.Text = " " Or Text2.Text = " " Or Text3.Text = " " Or Text4.Text = " " Or Text5.Text = " ") Then
        MsgBox "Field should not be left Blank"
        Exit Sub
    End If
    a = Val(Text3) + Val(Text4) + Val(Text5)
    Label10.Caption = a
    b = a / 3
    Label11.Caption = b
    If (b < 40) Then
        Label12.Caption = "FAIL"
    ElseIf (b < 45) Then
        Label12.Caption = "THIRD"
    ElseIf (b < 60) Then
        Label12.Caption = "SECOND"
    ElseIf (b < 75) Then
        Label12.Caption = "FIRST"
    Else
        Label12.Caption = "DISTINCTION"
    End If
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Text1.SetFocus
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Text1_keypress(keyAscii As Integer)
    If keyAscii = 13 And Text1.Text <> "" Then
        Text2.SetFocus
    ElseIf (keyAscii < 65 And keyAscii <> 8 And keyAscii <> 32) Or (keyAscii > 90 And keyAscii < 97) Or (keyAscii > 122) Then
        keyAscii = 0
        MsgBox "Enter Letter Only"
    End If
End Sub

Private Sub Text2_keypress(keyAscii As Integer)
    If keyAscii = 13 And Text2.Text <> "" Then
        Text3.SetFocus
    ElseIf (keyAscii < 65 And keyAscii <> 8 And keyAscii <> 32) Or (keyAscii > 90 And keyAscii < 97) Or (keyAscii > 122) Then
        keyAscii = 0
        MsgBox "Enter Letter Only"
    End If
End Sub

Private Sub text3_change()
    If Val(Text3.Text) > 100 Then
        MsgBox "Marks Range from 0 to 100"
        Text3.Text = ""
    End If
End Sub


Private Sub Text3_keypress(keyAscii As Integer)
    If keyAscii = 13 And Text3.Text <> "" Then
        Text4.SetFocus
    ElseIf (keyAscii < 48 And keyAscii <> 8) Or keyAscii > 57 Then
        keyAscii = 0
        MsgBox "Enter Digits Only"
    End If
End Sub


Private Sub text4_change()
    If Val(Text4.Text) > 100 Then
        MsgBox "Enter Digits Only"
    End If
End Sub

Private Sub text5_change()
    If Val(Text5.Text) > 100 Then
        MsgBox "Marks Range from 0 to 100"
        Text5.Text = ""
    End If
End Sub

Private Sub Text5_keypress(keyAscii As Integer)
    If keyAscii = 13 And Text5.Text <> "" Then
        cal.SetFocus
    ElseIf (keyAscii < 48 And keyAscii <> 8) Or keyAscii > 57 Then
        keyAscii = 0
        MsgBox "Enter Digits Only"
    End If
End Sub
