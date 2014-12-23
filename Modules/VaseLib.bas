Attribute VB_Name = "VaseLib"
'=======================
'--- Constants       ---
'=======================
Public Const METHOD_HEADER_PATTERN As String = _
    "Public Sub " & VaseConfig.TEST_METHOD_PATTERN

'=======================
'- Internal Functions  -
'=======================

'# This finds the modules that are deemed as test modules
Public Function FindTestModules(Book As Workbook) As Variant
    Dim Module As VBComponent, Modules As Variant, Index As Integer
    Modules = Array()
    Index = 0
    ReDim Modules(0 To Book.VBProject.VBComponents.Count)
    For Each Module In Book.VBProject.VBComponents
        If Module.Name Like VaseConfig.TEST_MODULE_PATTERN Then
            Set Modules(Index) = Module
            Index = Index + 1
        End If
    Next
    
    ' Fit array
    If Index = 0 Then
        Modules = Array()
    Else
        ReDim Preserve Modules(0 To Index - 1)
    End If
    
    FindTestModules = Modules
End Function

'# Finds the test methods to execute for a module
'@ Return: A zero-based string array of the method names to execute
Public Function FindTestMethods(Module As VBComponent) As Variant
    Dim Methods As Variant, Index As Integer, LineIndex As Integer, CodeLine As String
    Methods = Array()
    ReDim Methods(0 To Module.CodeModule.CountOfLines)
    
    For LineIndex = 1 To Module.CodeModule.CountOfLines
        CodeLine = Module.CodeModule.Lines(LineIndex, 1)
        If CodeLine Like METHOD_HEADER_PATTERN Then
            Dim LeftPos As Integer, RightPos As Integer
            LeftPos = InStr(CodeLine, "Sub") + 4
            RightPos = InStr(LeftPos, CodeLine, "(") - 1
            
            Methods(Index) = Mid(CodeLine, LeftPos, RightPos - LeftPos + 1)
            Index = Index + 1
        End If
    Next
    
    If Index = 0 Then
        Methods = Array()
    Else
        ReDim Preserve Methods(0 To Index - 1)
    End If
    FindTestMethods = Methods
End Function


'=======================
'-- Helper Functions  --
'=======================

'# Determines if a string is in an array using the like operator instead of equality
'@ Param: Patterns > An array of strings, not necessarily zero-based
'@ Return: True if the string matches any one of the patterns
Public Function InLike(Source As String, Patterns As Variant) As Boolean
    Dim Pattern As Variant
    InLike = False
    For Each Pattern In Patterns
        InLike = Source Like Pattern
        If InLike Then Exit For
    Next
End Function
