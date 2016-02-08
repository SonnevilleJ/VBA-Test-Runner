Attribute VB_Name = "TestModule"
Option Explicit

Public Sub RunAllTests()
    Call TestRunner.ExecuteTests
End Sub

Public Sub AssertAreEqual(expected As Integer, actual As Integer)
    Err.Clear
    Err.Source = "AssertAreEqual"
    If (expected <> actual) Then
        Err.Number = vbObjectError + 101
        Err.Description = "Actual: " + CStr(actual) + " differs from expected: " + CStr(expected) + "."
    End If
End Sub
