VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmch2 
   Caption         =   "Form 2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "Open"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdopen_Click()
    Dim filelocation As String
    CommonDialog1.ShowOpen
    filelocation = CommonDialog1.FileName
    If filelocation <> "" Then
        Open filelocation For Input As #1
        Do Until EOF(1)
            Input #1, Data
            Text1.Text = Text1.Text + Data + vbNewLine
        Loop
        Close #1
    Else
        MsgBox "No FileName selected ->->"
    End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

