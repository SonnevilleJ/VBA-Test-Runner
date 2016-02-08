Attribute VB_Name = "TestRunner"
Option Explicit

Const TestRunnerWorkbook As String = "VBA Unit Testing.xlsm"
Dim row As Integer

Public Sub ExecuteTests()
On Error GoTo errHandler
    Dim params() As Variant
    Dim sc As New ScriptControl
    sc.Language = "VBScript"
    
    sc.Eval "BadTest()"
    
    sc.AddCode "Public Sub BadTest()" & vbNewLine & "MsgBox 5/0" & vbNewLine & "End Sub"
    Call sc.Run("BadTest")
    
    MsgBox "shouldn't get here"
    Exit Sub

errHandler:
MsgBox "Error caught: " & Err.Number & ": " & Err.Description
Err.Clear
Resume Next

    Dim testModules As Collection
    Dim Workbook As Integer
    For Workbook = 1 To Workbooks.Count
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells.Clear
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(1, 1).Value = CStr(TimeValue(Now)) + ": Beginning test run..."
        
        row = 2
        If (Workbooks(Workbook).Name <> "VBA Unit Tests.xlsm") Then
            Set testModules = GetAllTestModules(Workbooks(Workbook).VBProject.VBComponents)
            
            Dim i As Integer
            For i = 1 To testModules.Count
                Call RunTestsInModule(Workbook, testModules(i))
            Next i
    
        End If
    Next Workbook
    
    Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(row, 1) = CStr(TimeValue(Now)) + ": Test run complete."
    Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Activate
End Sub

Private Sub RunTestsInModule(Workbook As Integer, TestModule As String)
    Dim testProcedures As Collection
    Set testProcedures = GetAllTestProceduresInModule(Workbook, TestModule)
    Dim i As Integer
    On Error GoTo TestFailed
    For i = 1 To testProcedures.Count
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(row, 1).Value = "Running..."
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(row, 2).Value = TestModule + "." + testProcedures(i)
        Call Workbooks(Workbook).Application.Run(testProcedures(i))
        GoTo TestSucceeded
        
TestFailed:
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(row, 1).Value = "Failed:"
        'Debug.Print "Test failed: " + testProcedures(i)
        Err.Clear
        GoTo NextTest
        
TestSucceeded:
        Workbooks(TestRunnerWorkbook).Worksheets("TestResults").Cells(row, 1).Value = "Passed:"
        'Debug.Print "Test succeeded: " + testProcedures(i)
        
NextTest:
        row = row + 1
        Next i
End Sub

Private Function GetAllTestModules(components As VBComponents) As VBA.Collection
    Dim comp As VBComponent
    Dim results As New Collection
    For Each comp In components
        If (comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_Document) Then
            If (InStrRev(comp.Name, "_Test")) Then
                results.Add (comp.Name)
            End If
        End If
    Next comp
    Set GetAllTestModules = results
End Function

Private Function GetAllTestProceduresInModule(Workbook As Integer, moduleName As String) As VBA.Collection
    Dim results As New Collection
    Dim LineNum As Long
    Dim NumLines As Long
    Dim module As CodeModule
    Set module = GetCodeModule(Workbook, moduleName)
    Dim procedure As String
    Dim prefix As String
    Dim paramsStart As Integer, paramsEnd As Integer
    With module
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            procedure = .ProcOfLine(LineNum, vbext_pk_Proc)
            prefix = "Public Sub " + procedure + "()"
            If InStr(1, .Lines(.ProcBodyLine(procedure, vbext_pk_Proc), 1), prefix, vbBinaryCompare) <> 0 Then
                results.Add procedure
            End If
            LineNum = .ProcStartLine(procedure, vbext_pk_Proc) + .ProcCountLines(procedure, vbext_pk_Proc) + 1
        Loop
    End With
    Set GetAllTestProceduresInModule = results
End Function

Private Function GetCodeModule(Workbook As Integer, module As String) As CodeModule
    Dim result As CodeModule
    Dim components As VBComponents
    Set components = Workbooks(Workbook).VBProject.VBComponents
    Dim i As Integer
    For i = 1 To components.Count
        If (components(i).Name = module) Then
            Set result = components(i).CodeModule
            Exit For
        End If
    Next
    Set GetCodeModule = result
End Function
