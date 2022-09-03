# Visual-Basic-Programs


### **1. Write a VB program to design a simple calculator to perform addition,subtraction,multiplication and division(Use functions for the calculations).**
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
### **2. Design a User Interface (UI) to accept the student details such as name,department and total marks.Validate the input data and calculate the percentage nd division**
![Screenshot (3)](https://user-images.githubusercontent.com/74803670/177051946-ba4cd5a8-d301-4364-9abb-d421b691e55d.png)
```
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
```
---
---



