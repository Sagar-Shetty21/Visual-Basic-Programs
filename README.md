# Visual-Basic-Programs


## **1. Write a VB program to design a simple calculator to perform addition,subtraction,multiplication and division(Use functions for the calculations).**
![Screenshot (2)](https://user-images.githubusercontent.com/74803670/177051678-607fe38e-59b2-4d31-ac91-5f5c04559780.png)

```
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
```
---



---
## **2. Design a VB application which has MDI and child forms. Create a menu having the items such as File(New,Open),Format(Font,Regular,Bold,Italic) and Exit in the MDI form. Also create a textbox and use a Common Dialog Box for changing the font, forecolor and back color of the text box.**

### **MDI Form**
![Screenshot (23)](https://user-images.githubusercontent.com/74803670/188280025-7b164703-7a54-414d-ab31-39615045cdd2.png)
![Screenshot (24)](https://user-images.githubusercontent.com/74803670/188280062-9be6bfc3-a29f-488b-bfda-7b7c1b937485.png)

```
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
```

### **Child Forms**

** Form 1 **
```
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
```

** Form 2 **
```
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
```

** Form 3 **
```
Private Sub cmdexit_Click()
    Unload Me
End Sub
```
---
---

## ** 3.  **



