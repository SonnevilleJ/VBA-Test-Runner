Attribute VB_Name = "Module1_Test"
Option Explicit

Public Sub FirstUnitTest()
    Const expected As Integer = 5
    
    Dim actual As Integer
    actual = Add(2, 3)
    
    Call AssertAreEqual(expected, actual)
End Sub

Public Sub SecondTest()
    Const expected As Integer = 5
    
    Dim actual As Integer
    actual = Add(2, 2)
    
    Call AssertAreEqual(expected, actual)
End Sub

Public Sub BadTest()
    Dim stupid As Integer
    stupid = 1 / 0
End Sub
