VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmch1 
   Caption         =   "Form 1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
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
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdsave_Click()
    Dim filelocation As String
    If Text1.Text <> "" Then
        CommonDialog1.ShowSave
        filelocation = CommonDialog1.FileName
        If filelocation <> "" Then
            Open filelocation For Append As #1
            Print #1, Text1.Text
            Close #1
        Else
            MsgBox "File name not specified. File cannot be saved"
        End If
    Else
        MsgBox "Text box is empty"
    End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub


