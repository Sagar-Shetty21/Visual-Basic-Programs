VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Vending Machine Application"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msflex1 
      Height          =   1695
      Left            =   600
      TabIndex        =   14
      Top             =   4440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
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
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmddespense 
      Caption         =   "DESPENSE SNACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtbill 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtqty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton optsnack4 
      Caption         =   "chocolate"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton optsnack3 
      Caption         =   "samosa"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton optsnack2 
      Caption         =   "almond"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton optsnack1 
      Caption         =   "Pepsi"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frame1 
      Caption         =   "SELECT YOUR SNACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lbldisplay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL BILL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "DESPENDED SNACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "ENTER QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VENDING MACHINE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slno As Integer
Dim rno As Integer
Dim totalbill As Integer


Private Sub cmddespense_Click()
    Dim sname As String
    lbldisplay.Caption = ""
    If txtqty.Text = "" Or Val(txtqty.Text) <= 0 Then
        MsgBox "Quantity cannot be blanck or 0 or negetive"
        Exit Sub
    End If
    If optsnack1.Value = False And optsnack2.Value = False And optsnack3.Value = False And optsnack4.Value = False Then
        MsgBox "Please select atleast one snack"
        Exit Sub
    End If
    If optsnack1.Value = True Then
        sname = optsnack1.Caption
        sprice = 20
    ElseIf optsnack2.Value = True Then
        sname = optsnack2.Caption
        sprice = 25
    ElseIf optsnack3.Value = True Then
        sname = optsnack3.Caption
        sprice = 25
    ElseIf optsnack4.Value = True Then
        sname = optsnack4.Caption
        sprice = 25
    End If
    lbldisplay.Caption = sname
    msflex1.Rows = msflex1.Rows + 1
    msflex1.TextMatrix(rno, 0) = slno
    msflex1.TextMatrix(rno, 1) = sname
    msflex1.TextMatrix(rno, 2) = sprice
    msflex1.TextMatrix(rno, 3) = txtqty.Text
    msflex1.TextMatrix(rno, 4) = Val(txtqty.Text) * sprice
    totalbill = totalbill + Val(msflex1.TextMatrix(rno, 4))
    txtbill = totalbill
    slno = slno + 1
    rno = rno + 1
    txtqty = ""
End Sub



Private Sub cmdexit_Click()
    MsgBox "Total Bill Amount is " + txtbill.Text
    End
End Sub

Private Sub form_load()
    rno = 1
    slno = 1
    msflex1.Cols = 5
    msflex1.ColWidth(0) = 800
    msflex1.ColWidth(1) = 3000
    msflex1.ColWidth(2) = 1200
    msflex1.ColWidth(3) = 1200
    msflex1.ColWidth(4) = 1200
    optsnack1.Caption = "pepsi"
    optsnack2.Caption = "almond"
    optsnack3.Caption = "samosa"
    optsnack4.Caption = "chocolate"
End Sub
