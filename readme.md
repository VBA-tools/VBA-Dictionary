# VBA-Dictionary

VBA-Dictionary is a drop-in replacement for the useful and powerful `Scripting.Dictionary` on Mac. It is designed to be a precise replacement to `Scripting.Dictionary` including `Items` as the default property (`Dict("A") = Dict.Items("A")`), matching error codes, and matching methods and properties. If you find any implementation differences between `Scripting.Dictionary` and VBA-Dictionary, please [create an issue](https://github.com/timhall/VBA-Dictionary/issues/new).

## Example

```VB
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

Dict.Key("B") = "C"
Dict.Exists "B" ' -> False
Dict("C")("Inner") ' -> "Value"

Dict.CompareMode = CompareMethods.BinaryCompare
' -> Throws 5 (Can't change CompareMode when there are items in the Dictionary)

Dict.Remove "B"
' -> Throws 32811: Application-defined or object-defined error

Dict.Remove "A"
Dict.RemoveAll

Dict.Exists "A" ' -> False
Dict("C") ' -> Empty
UBound(Dict.Keys) ' -> -1
UBound(Dict.Items) ' -> -1
```

### Release Notes

#### 1.0.0

- Complete replacement for `Scripting.Dictionary`
- __1.0.1__ Allow Variant keys and handle empty Dictionary
- __1.0.2__ Fix replace with single item in Dictionary
- __1.0.3__ Documentation fixes
