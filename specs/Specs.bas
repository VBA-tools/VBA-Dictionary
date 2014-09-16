Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    InlineRunner.RunSuite RunSpecs(True)
    InlineRunner.RunSuite RunSpecs(False)
End Function

Public Function RunSpecs(Optional UseNative As Boolean = False) As SpecSuite
    Dim Specs As New SpecSuite
    If UseNative Then
        Specs.Description = "Scripting.Dictionary"
    Else
        Specs.Description = "VBA-Dictionary"
    End If
    
    Dim Dict As Object
    Dim Items As Variant
    Dim Keys As Variant
    Dim Key As Variant
    Dim Item As Variant
    
    ' Properties
    ' ------------------------- '
    With Specs.It("should get count of items")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        .Expect(Dict.Count).ToEqual 3
        
        Dict.Remove "C"
        .Expect(Dict.Count).ToEqual 2
    End With
    
    With Specs.It("should get item by key")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        .Expect(Dict.Item("B")).ToEqual 3.14
        .Expect(Dict.Item("D")).ToBeEmpty
        .Expect(Dict("B")).ToEqual 3.14
        .Expect(Dict("D")).ToBeEmpty
    End With
    
    With Specs.It("should let/set item by key")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        Dict.Item("D") = True
        Dict("C") = "DEF"
        
        Set Dict.Item("B") = CreateDictionary(UseNative)
        Dict.Item("B").Add "key", "B"
        Set Dict("A") = CreateDictionary(UseNative)
        Dict("A").Add "key", "A"
        
        .Expect(Dict.Item("A")("key")).ToEqual "A"
        .Expect(Dict.Item("B")("key")).ToEqual "B"
        .Expect(Dict.Item("C")).ToEqual "DEF"
        .Expect(Dict.Item("D")).ToEqual True
    End With
    
    With Specs.It("should change key")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        Dict.Key("B") = "PI"
        .Expect(Dict("PI")).ToEqual 3.14
    End With
    
    With Specs.It("should use CompareMode")
        Set Dict = CreateDictionary(UseNative)
        Dict.CompareMode = 0
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        .Expect(Dict.Exists("a")).ToEqual False
        
        Set Dict = CreateDictionary(UseNative)
        Dict.CompareMode = 1
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        .Expect(Dict.Exists("a")).ToEqual True
    End With
    
    With Specs.It("should allow Variant for key")
        Set Dict = CreateDictionary(UseNative)
        
        Key = "A"
        Dict(Key) = 123
        .Expect(Dict(Key)).ToEqual 123
        
        Key = "B"
        Set Dict(Key) = CreateDictionary(UseNative)
        .Expect(Dict(Key).Count).ToEqual 0
    End With
    
    ' Methods
    ' ------------------------- '
    With Specs.It("should add an item")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        Dict.Add "E", Array(1, 2, 3)
        Dict.Add "F", Dict
        
        .Expect(Dict("A")).ToEqual 123
        .Expect(Dict("B")).ToEqual 3.14
        .Expect(Dict("C")).ToEqual "ABC"
        .Expect(Dict("D")).ToEqual True
        .Expect(Dict("E")(1)).ToEqual 2
        .Expect(Dict("F")("C")).ToEqual "ABC"
    End With
    
    With Specs.It("should check if an item exists")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "Exists", 123
        .Expect(Dict.Exists("Exists")).ToEqual True
        .Expect(Dict.Exists("Doesn't Exist")).ToEqual False
    End With
    
    With Specs.It("should get an array of all items")
        Set Dict = CreateDictionary(UseNative)
        
        .Expect(UBound(Dict.Items)).ToEqual -1
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Items = Dict.Items
        .Expect(UBound(Items)).ToEqual 3
        .Expect(Items(0)).ToEqual 123
        .Expect(Items(3)).ToEqual True
    End With
    
    With Specs.It("should get an array of all keys")
        Set Dict = CreateDictionary(UseNative)
        
        .Expect(UBound(Dict.Keys)).ToEqual -1
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Keys = Dict.Keys
        .Expect(UBound(Keys)).ToEqual 3
        .Expect(Keys(0)).ToEqual "A"
        .Expect(Keys(3)).ToEqual "D"
    End With
    
    With Specs.It("should remove item")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        .Expect(Dict.Count).ToEqual 4
        
        Dict.Remove "C"
                
        .Expect(Dict.Count).ToEqual 3
    End With
    
    With Specs.It("should remove all items")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        .Expect(Dict.Count).ToEqual 4
        
        Dict.RemoveAll
        
        .Expect(Dict.Count).ToEqual 0
    End With
    
    ' Other
    ' ------------------------- '
    With Specs.It("should For Each over keys")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Set Keys = New Collection
        For Each Key In Dict.Keys
            Keys.Add Key
        Next Key
        
        .Expect(Keys.Count).ToEqual 4
        .Expect(Keys(1)).ToEqual "A"
        .Expect(Keys(4)).ToEqual "D"
    End With
    
    With Specs.It("should For Each over items")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Set Items = New Collection
        For Each Item In Dict.Items
            Items.Add Item
        Next Item
        
        .Expect(Items.Count).ToEqual 4
        .Expect(Items(1)).ToEqual 123
        .Expect(Items(4)).ToEqual True
    End With
    
    Set RunSpecs = Specs
End Function

Public Function CreateDictionary(Optional UseNative As Boolean = False) As Object
    If UseNative Then
        Set CreateDictionary = New Scripting.Dictionary
    Else
        Set CreateDictionary = New VBAProject.Dictionary
    End If
End Function
