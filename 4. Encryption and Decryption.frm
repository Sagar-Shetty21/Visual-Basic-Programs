VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encryption and Decryption"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4440
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
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
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ENCRYPTED STRING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ENTER THE TEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "encryption and decryption"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function encryptdecrypt(ByVal strval, ByVal blndec As Boolean) As String
    Dim int1, int2, int3 As Integer
    str1 = strval
    int1 = Len(str1)
    int2 = 1
    Randomize
    Do While int2 <> (int1 + 1)
        str2 = Mid(str1, int2, 1)
        If blndec = True Then
            int3 = Asc(str2) - 3
        Else
            int3 = Asc(str2) + 3
        End If
        str3 = Chr(int3)
        str4 = str4 & str3
        int2 = int2 + 1
    Loop
    encryptdecrypt = str4
End Function


Private Sub Command1_Click()
    x = encryptdecrypt(Text1.Text, True)
    MsgBox (x)
    Text2.Text = x
End Sub


Private Sub Command2_Click()
    x = encryptdecrypt(Text2.Text, False)
    MsgBox (x)
End Sub

Private Sub Command3_Click()
    Text1.Text = " "
    Text2.Text = " "
End Sub

Private Sub command4_click()
    End
End Sub
