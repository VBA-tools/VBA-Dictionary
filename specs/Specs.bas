Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-Dictionary"
    
    InlineRunner.RunSuite Specs
End Function
