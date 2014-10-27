Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    #If Mac Then
        ' Mac
        InlineRunner.RunSuite AllSpecs(UseNative:=False)
        SpeedTest CompareToNative:=False
    #Else
        ' Windows
        InlineRunner.RunSuite AllSpecs(UseNative:=True)
        InlineRunner.RunSuite AllSpecs(UseNative:=False)
        SpeedTest CompareToNative:=True
    #End If
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite AllSpecs(UseNative:=False)
End Sub

Public Function AllSpecs(Optional UseNative As Boolean = False) As SpecSuite
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
    
    On Error Resume Next
    
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
        Dict("A") = 456
        
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
        Dict("a") = 456
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        .Expect(Dict("A")).ToEqual 123
        .Expect(Dict("a")).ToEqual 456
        
        Set Dict = CreateDictionary(UseNative)
        Dict.CompareMode = 1
        
        Dict.Add "A", 123
        Dict("a") = 456
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        
        .Expect(Dict("A")).ToEqual 456
        .Expect(Dict("a")).ToEqual 456
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
        
        .Expect(Dict.Items).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Items = Dict.Items
        .Expect(UBound(Items)).ToEqual 3
        .Expect(Items(0)).ToEqual 123
        .Expect(Items(3)).ToEqual True
        
        Dict.Remove "A"
        Dict.Remove "B"
        Dict.Remove "C"
        Dict.Remove "D"
        .Expect(Dict.Items).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
    End With
    
    With Specs.It("should get an array of all keys")
        Set Dict = CreateDictionary(UseNative)
        
        .Expect(Dict.Keys).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
        
        Dict.Add "A", 123
        Dict.Add "B", 3.14
        Dict.Add "C", "ABC"
        Dict.Add "D", True
        
        Keys = Dict.Keys
        .Expect(UBound(Keys)).ToEqual 3
        .Expect(Keys(0)).ToEqual "A"
        .Expect(Keys(3)).ToEqual "D"
        
        Dict.RemoveAll
        .Expect(Dict.Keys).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
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
        
        Set Keys = New Collection
        For Each Key In Dict.Keys
            Keys.Add Key
        Next Key
        
        .Expect(Keys.Count).ToEqual 0
        
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
        Err.Clear
    End With
    
    With Specs.It("should For Each over items")
        Set Dict = CreateDictionary(UseNative)
        
        Set Items = New Collection
        For Each Item In Dict.Items
            Items.Add Item
        Next Item
        
        .Expect(Items.Count).ToEqual 0
        
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
        Err.Clear
    End With
    
    With Specs.It("should have UBound of -1 for empty Keys and Items")
        Set Dict = CreateDictionary(UseNative)
        
        .Expect(UBound(Dict.Keys)).ToEqual -1
        .Expect(UBound(Dict.Items)).ToEqual -1
        .Expect(Err.Number).ToEqual 0
        Err.Clear
    End With
    
    ' Errors
    ' ------------------------- '
    Err.Clear
    With Specs.It("should throw 5 when changing CompareMode with items in Dictionary")
        Set Dict = CreateDictionary(UseNative)
        Dict.Add "A", 123
        
        Dict.CompareMode = vbTextCompare
        
        .Expect(Err.Number).ToEqual 5
        Err.Clear
    End With
    
    With Specs.It("should throw 457 on Add if key exists")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Add "A", 123
        Dict.Add "A", 456
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
        
        Dict.RemoveAll
        Dict.Add "A", 123
        Dict.Add "a", 456
        
        .Expect(Err.Number).ToEqual 0
        Err.Clear
        
        Dict.RemoveAll
        Dict.CompareMode = vbTextCompare
        Dict.Add "A", 123
        Dict.Add "a", 456
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
    End With
    
    With Specs.It("should throw 32811 on Remove if key doesn't exist")
        Set Dict = CreateDictionary(UseNative)
        
        Dict.Remove "A"
        
        .Expect(Err.Number).ToEqual 32811
        Err.Clear
    End With
    
    On Error GoTo 0
    
    Set AllSpecs = Specs
End Function

Public Sub SpeedTest(Optional CompareToNative As Boolean = False)
    Dim Counts As Variant
    Counts = Array(5000, 5000, 5000, 5000, 7500, 7500, 7500, 7500)
    
    Dim Baseline As Collection
    If CompareToNative Then
        Set Baseline = RunSpeedTest(Counts, True)
    End If
    
    Dim Results As Collection
    Set Results = RunSpeedTest(Counts, False)
    
    Debug.Print vbNewLine & "SpeedTest Results:" & vbNewLine
    PrintResults "Add", Baseline, Results, 0
    PrintResults "Iterate", Baseline, Results, 1
End Sub

Public Sub PrintResults(Test As String, Baseline As Collection, Results As Collection, Index As Integer)
    Dim BaselineAvg As Single
    Dim ResultsAvg As Single
    Dim i As Integer
    
    If Not Baseline Is Nothing Then
        For i = 1 To Baseline.Count
            BaselineAvg = BaselineAvg + Baseline(i)(Index)
        Next i
        BaselineAvg = BaselineAvg / Baseline.Count
    End If
    
    For i = 1 To Results.Count
        ResultsAvg = ResultsAvg + Results(i)(Index)
    Next i
    ResultsAvg = ResultsAvg / Results.Count
    
    Dim Result As String
    Result = Test & ": " & Format(Round(ResultsAvg, 0), "#,##0") & " ops./s"
    
    If Not Baseline Is Nothing Then
        Result = Result & " vs. " & Format(Round(BaselineAvg, 0), "#,##0") & " ops./s "
    
        If ResultsAvg < BaselineAvg Then
            Result = Result & Format(Round(BaselineAvg / ResultsAvg, 0), "#,##0") & "x slower"
        ElseIf BaselineAvg > ResultsAvg Then
            Result = Result & Format(Round(ResultsAvg / BaselineAvg, 0), "#,##0") & "x faster"
        End If
    End If
    Result = Result
    
    If Results.Count > 1 Then
        Result = Result & vbNewLine
        For i = 1 To Results.Count
            Result = Result & "  " & Format(Round(Results(i)(Index), 0), "#,##0")
            
            If Not Baseline Is Nothing Then
                Result = Result & " vs. " & Format(Round(Baseline(i)(Index), 0), "#,##0")
            End If
            
            Result = Result & vbNewLine
        Next i
    End If
    
    Debug.Print Result
End Sub

Public Function RunSpeedTest(Counts As Variant, Optional UseNative As Boolean = False) As Collection
    Dim Results As New Collection
    Dim CountIndex As Integer
    Dim Dict As Object
    Dim i As Long
    Dim AddResult As Double
    Dim Key As Variant
    Dim Value As Variant
    Dim IterateResult As Double
    Dim Timer As New PreciseTimer
    
    For CountIndex = LBound(Counts) To UBound(Counts)
        Timer.StartTimer
    
        Set Dict = CreateDictionary(UseNative)
        For i = 1 To Counts(CountIndex)
            Dict.Add CStr(i), i
        Next i
        
        ' Convert to seconds
        AddResult = Timer.TimeElapsed / 1000#
        
        ' Convert to ops./s
        If AddResult > 0 Then
            AddResult = Counts(CountIndex) / AddResult
        Else
            ' Due to single precision, timer resolution is 0.01 ms, set to 0.005 ms
            AddResult = Counts(CountIndex) / 0.005
        End If
        
        Timer.StartTimer
        
        For Each Key In Dict.Keys
            Value = Dict.Item(Key)
        Next Key
        
        ' Convert to seconds
        IterateResult = Timer.TimeElapsed / 1000#
        
        ' Convert to ops./s
        If IterateResult > 0 Then
            IterateResult = Counts(CountIndex) / IterateResult
        Else
            ' Due to single precision, timer resolution is 0.01 ms, set to 0.005 ms
            IterateResult = Counts(CountIndex) / 0.005
        End If
        
        Results.Add Array(AddResult, IterateResult)
    Next CountIndex
    
    Set RunSpeedTest = Results
End Function

Public Function CreateDictionary(Optional UseNative As Boolean = False) As Object
    If UseNative Then
        Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Else
        Set CreateDictionary = New Dictionary
    End If
End Function

Public Function ToBeAnEmptyArray(Actual As Variant) As Variant
    Dim UpperBound As Long

    Err.Clear
    On Error Resume Next
    
    ' First, make sure it's an array
    If IsArray(Actual) = False Then
        ' we weren't passed an array, return True
        ToBeAnEmptyArray = True
    Else
        ' Attempt to get the UBound of the array. If the array is
        ' unallocated, an error will occur.
        UpperBound = UBound(Actual, 1)
        If (Err.Number <> 0) Then
            ToBeAnEmptyArray = True
        Else
            ' Check for case of -1 UpperBound (Scripting.Dictionary.Keys/Items)
            Err.Clear
            If LBound(Actual) > UpperBound Then
                ToBeAnEmptyArray = True
            Else
                ToBeAnEmptyArray = False
            End If
        End If
    End If
End Function
