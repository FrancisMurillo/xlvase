Attribute VB_Name = "VaseAssert"
'# This global variable determines if an assert method failed or passed
Private gAssertion As Boolean
Private gFirstFailed As String

Public Property Get TestResult() As Boolean
    TestResult = gAssertion
End Property
Public Property Get FirstFailedTestMethod() As String
    FirstFailedTestMethod = gFirstFailed
End Property

'=======================
'--- Assertion Tools ---
'=======================

'# Sets the Assertion globals for use
Public Sub InitAssert()
    gAssertion = True
    gFirstFailed = ""
End Sub


'# Base Assert Method
Public Sub Assert_(Cond As Boolean, _
        Optional Message As String = "", _
        Optional AssertName As String = "Assert") ' Name to avoid Debug.Assert conflict or confusion
    gAssertion = gAssertion And Cond ' Update assertion variable
    If Not Cond And gFirstFailed = "" Then ' Log the first fail condition for logging
        gFirstFailed = AssertName
    End If
End Sub

'# Assert if condition is true
Public Sub AssertTrue(Cond As Boolean, Optional Message As String = "")
    Assert_ Cond, Message:=Message, AssertName:="AssertTrue"
End Sub

'# Assert condition is false
Public Sub AssertFalse(Cond As Boolean, Optional Message As String = "")
    Assert_ Not Cond, Message:=Message, AssertName:="AssertFalse"
End Sub

'# Assert two variables are equal
Public Sub AssertEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertEqual"
End Sub

Private Function Equal_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Equal_ = False
    Equal_ = (LeftVal = RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function
