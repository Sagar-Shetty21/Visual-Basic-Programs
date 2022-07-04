VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Log-in Form"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc OLEDB 
      Height          =   375
      Left            =   2880
      Top             =   3000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2880
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
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
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOGIN"
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
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Line Line6 
      X1              =   840
      X2              =   6240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   6240
      X2              =   6240
      Y1              =   240
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   6240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   840
      Y1              =   240
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   6240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ACCOUNT LOGIN"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.connection

Private Sub cmdexit_Click()
    MsgBox "Do you really want to Exit", vbInformation + vbOKOnly, "login"
    If vbOK Then
        End
    End If
End Sub

Private Sub cmdlogin_Click()
    If Text1.Text = "" Then
        MsgBox "enter the username", vbInformation + vbOKOnly, "login"
        Text1.SetFocus
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "enter the password", vbInformation + vbOKOnly, "login"
        Text2.SetFocus
        Exit Sub
    End If
    If Text1.Text <> "" And Text2.Text <> "" Then
        If rs.state = 1 Then
            rs.Close
        Else
            rs.Open "select* from login where username =""&text1&""and password="" &text2 &"",conn,adopendynamic,adlockoptimistic,adcmdtext"
    End If
    If rs.EOF = True Then
        MsgBox "invalid username and password", vbCritical + vbOKOnly, "login"
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
    Else
        MsgBox "username and password correct", vbInformation + vbOKOnly, "login"
    End If
End Sub

Private Sub form_load()
    conn.Open "provider=microsoft.jet.OLEDB.4.0;data source=" & App.Path & "\Account.mdb;persist security info=false"
End Sub
