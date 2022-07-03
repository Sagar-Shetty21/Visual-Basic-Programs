VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000006&
   Caption         =   "MDI Form"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "Format"
      Begin VB.Menu mnufont 
         Caption         =   "Font"
         Begin VB.Menu mnuarial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnugaramond 
            Caption         =   "Garamond"
         End
         Begin VB.Menu mnuimpact 
            Caption         =   "Impact"
         End
         Begin VB.Menu mnulucida 
            Caption         =   "Lucida"
         End
      End
      Begin VB.Menu mnuregular 
         Caption         =   "Regular"
      End
      Begin VB.Menu mnubold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuitalic 
         Caption         =   "Italic"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnunew_Click()
    frmch1.Show
End Sub


Private Sub mnuopen_Click()
    frmch2.Show
End Sub

Private Sub mnuformat_Click()
    frmch3.Show
End Sub

Private Sub mnuarial_Click()
    frmch3.txtdemo.FontName = "Arial"
End Sub

Private Sub mnugaramond_Click()
    frmch3.txtdemo.FontName = "garamond"
End Sub

Private Sub mnuimpact_Click()
    frmch3.txtdemo.FontName = "impact"
End Sub

Private Sub mnulucida_Click()
    frmch3.txtdemo.FontName = "lucida sans"
End Sub


Private Sub mnubold_Click()
    frmch3.txtdemo.FontBold = True
End Sub

Private Sub mnuitalic_Click()
    frmch3.txtdemo.FontItalic = True
End Sub

Private Sub mnuregular_Click()
    frmch3.txtdemo.FontBold = False
    frmch3.txtdemo.FontItalic = False
End Sub


Private Sub mnuexit_Click()
    Unload Me
End Sub

