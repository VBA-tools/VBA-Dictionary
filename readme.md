# VBA-Dictionary

VBA-Dictionary is a drop-in replacement for the useful and powerful `Scripting.Dictionary` so that it can be used with both Mac and Windows. It is designed to be a precise replacement to `Scripting.Dictionary` including `Item` as the default property (`Dict("A") = Dict.Item("A")`), matching error codes, and matching methods and properties. If you find any implementation differences between `Scripting.Dictionary` and VBA-Dictionary, please [create an issue](https://github.com/timhall/VBA-Dictionary/issues/new).

## Example

```VB.net
' (Works exactly like Scripting.Dictionary)
Dim Dict As New Dictionary
Dict.CompareMode = CompareMethods.TextCompare

Dict("A") ' -> Empty
Dict("A") = 123
Dict("A") ' -> = Dict.Item("A") = 123
Dict.Exists "A" ' -> True

Dict.Add "A", 456 
' -> Throws 457: This key is already associated with an element of this collection

' Both Set and Let work
Set Dict("B") = New Dictionary
Dict("B").Add "Inner", "Value"
Dict("B")("Inner") ' -> "Value"

UBound(Dict.Keys) ' -> 1
UBound(Dict.Items) ' -> 1

' Rename key
Dict.Key("B") = "C"
Dict.Exists "B" ' -> False
Dict("C")("Inner") ' -> "Value"

' Trying to remove non-existant key throws 32811
Dict.Remove "B"
' -> Throws 32811: Application-defined or object-defined error

' Trying to change CompareMode when there are items in the Dictionary throws 5
Dict.CompareMode = CompareMethods.BinaryCompare
' -> Throws 5: Invalid procedure call or argument

Dict.Remove "A"
Dict.RemoveAll

Dict.Exists "A" ' -> False
Dict("C") ' -> Empty
```

### Release Notes

#### 1.2.0

- Improve compatibility for empty Dictionary (`UBound` for empty `Keys` and `Items` is `-1` and can `For Each` over empty `Keys` and `Items`, matching `Scripting.Dictionary`)

#### 1.1.0

- Use compiler statements to use Scripting.Dictionary internally if available (improves Windows performance by ~3x)
- __1.1.1__ Make VBA-Dictionary instancing Public Not Creatable

#### 1.0.0

Initial release of VBA-Dictionary

- Exactly matches `Scripting.Dictionary` behavior (Methods/Properties, return types, errors thrown, etc.)
- Windows and Mac support (tested in Excel 2013 32-bit Windows and Excel 2011 Mac)
