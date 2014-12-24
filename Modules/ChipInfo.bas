Attribute VB_Name = "ChipInfo"
Public References As Variant
Public Modules As Variant

Public Sub Initialize()
    References = Array( _
        "Microsoft Visual Basic for Applications Extensibility *")
    Modules = Array( _
        "Vase", "VaseLib", "VaseAssert", "VaseConfig")
End Sub
