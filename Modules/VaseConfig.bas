Attribute VB_Name = "VaseConfig"
'=======================
'--- Configurations  ---
'=======================
'# These are the patterns for discovering which methods and modules are to be tested
Public Const TEST_MODULE_PATTERN As String = "Test*"
Public Const TEST_METHOD_PATTERN As String = "Test*"

'# Delimiters for string splitting
Public Const Delimiter As String = ";"

'# These are the names of the Setup and Teardown method
Public Const TEST_SETUP_METHOD_NAME As String = "Setup"
Public Const TEST_TEARDOWN_METHOD_NAME As String = "Teardown"
