# Visual-Basic-Programs


### **1. Write a VB program to design a simple calculator to perform addition,subtraction,multiplication and division(Use functions for the calculations).**
```Simple_Calculator.frm
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
