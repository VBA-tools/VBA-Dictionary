VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Dictionary v1.0.4
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Drop-in replacement for Scripting.Dictionary on Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' Store 0: Index, 1: Key, 2: Value
Private pKeyValues As Collection
Private pCompareMode As CompareMethod

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
    CompareMode = pCompareMode
End Property
Public Property Let CompareMode(Value As CompareMethod)
    If Me.Count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        Err.Raise 5 ' Invalid procedure call or argument
    End If
    
    pCompareMode = Value
End Property

Public Property Get Count() As Long
    Count = pKeyValues.Count
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(Key)
    
    If Not IsEmpty(KeyValue) Then
        If IsObject(KeyValue(2)) Then
            Set Item = KeyValue(2)
        Else
            Item = KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
End Property
Public Property Let Item(Key As Variant, Value As Variant)
    If Me.Exists(Key) Then
        ReplaceKeyValue GetKeyValue(Key), Key, Value
    Else
        AddKeyValue Key, Value
    End If
End Property
Public Property Set Item(Key As Variant, Value As Variant)
    If Me.Exists(Key) Then
        ReplaceKeyValue GetKeyValue(Key), Key, Value
    Else
        AddKeyValue Key, Value
    End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(Previous)
    
    If Not IsEmpty(KeyValue) Then
        ReplaceKeyValue KeyValue, Updated, KeyValue(2)
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(Key As Variant, Item As Variant)
    If Not Me.Exists(Key) Then
        AddKeyValue Key, Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
    Exists = Not IsEmpty(GetKeyValue(Key))
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
    Dim Arr As Variant
    
    If pKeyValues.Count > 0 Then
        Dim i As Long
        ReDim Arr(pKeyValues.Count - 1)
        For i = 1 To pKeyValues.Count
            If IsObject(pKeyValues(i)(2)) Then
                Set Arr(i - 1) = pKeyValues(i)(2)
            Else
                Arr(i - 1) = pKeyValues(i)(2)
            End If
        Next i
    Else
        Arr = Array()
    End If
    
    Items = Arr
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
    Dim Arr As Variant
    
    If pKeyValues.Count > 0 Then
        Dim i As Long
        ReDim Arr(pKeyValues.Count - 1)
        For i = 1 To pKeyValues.Count
            Arr(i - 1) = pKeyValues(i)(1)
        Next i
    Else
        Arr = Array()
    End If
    
    Keys = Arr
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(Key)
    
    If Not IsEmpty(KeyValue) Then
        RemoveKeyValue KeyValue
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
    Set pKeyValues = New Collection
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function GetKeyValue(Key As Variant) As Variant
    On Error Resume Next
    GetKeyValue = pKeyValues(GetFormattedKey(Key))
End Function

Private Sub AddKeyValue(Key As Variant, Value As Variant)
    Dim FormattedKey As String
    FormattedKey = GetFormattedKey(Key)
    pKeyValues.Add Array(pKeyValues.Count + 1, FormattedKey, Value), FormattedKey
End Sub

Private Sub ReplaceKeyValue(KeyValue As Variant, Key As Variant, Value As Variant)
    ' Remove previous KeyValue
    RemoveKeyValue KeyValue
    
    ' Add new key and value
    Dim Index As Long
    Dim NewKeyValue As Variant
    Index = KeyValue(0)
    NewKeyValue = Array(Index, GetFormattedKey(Key), Value)
    
    ' Add updated key value back to KeyValues
    If pKeyValues.Count = 0 Then
        pKeyValues.Add NewKeyValue, CStr(Key)
    ElseIf Index > pKeyValues.Count Then
        pKeyValues.Add NewKeyValue, CStr(Key), After:=Index - 1
    Else
        pKeyValues.Add NewKeyValue, CStr(Key), Before:=Index
    End If
End Sub

Private Sub RemoveKeyValue(KeyValue As Variant)
    pKeyValues.Remove KeyValue(0)
End Sub

Private Function GetFormattedKey(Key As Variant) As String
    GetFormattedKey = CStr(Key)
    If Me.CompareMode = CompareMethod.TextCompare Then
        GetFormattedKey = UCase(GetFormattedKey)
    ElseIf Me.CompareMode = CompareMethod.BinaryCompare Then
        ' Collection does not have method of setting key comparison
        ' So case-sensitive keys aren't supported by default
        ' -> Approach: Append lowercase characters to original key
        '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
        Dim Lowercase As String
        Lowercase = ""
        
        Dim i As Integer
        Dim Ascii As Integer
        Dim Char As String
        For i = 1 To Len(GetFormattedKey)
            Char = Mid$(GetFormattedKey, i, 1)
            Ascii = Asc(Char)
            If Ascii >= 97 And Ascii <= 122 Then
                Lowercase = Lowercase & Char
            End If
        Next i
        
        If Lowercase <> "" Then
            GetFormattedKey = GetFormattedKey & "__" & Lowercase
        End If
    End If
End Function

'Private Function GetKey(KeyValue As Variant) As String
'    If Not IsEmpty(KeyValue) Then
'        GetKey = KeyValue(0)
'    End If
'End Function
'Private Function GetValue(KeyValue As Variant) As Variant
'    If Not IsEmpty(KeyValue) Then
'        If IsObject(KeyValue(1)) Then
'            Set GetValue = KeyValue(1)
'        Else
'            GetValue = KeyValue(1)
'        End If
'    End If
'End Function
'Private Function GetIndex(KeyValue As Variant) As Long
'    If Not IsEmpty(KeyValue) Then
'        GetIndex = KeyValue(2)
'    End If
'End Function

'Private Function IndexOfKey(Key As Variant) As Long
'    Dim i As Long
'    Key = CStr(Key)
'
'    For i = 1 To pKeys.Count
'        If VBA.StrComp(pKeys(i), Key, Me.CompareMode) = 0 Then
'            IndexOfKey = i
'            Exit Function
'        End If
'    Next i
'
'    ' Not found
'    IndexOfKey = -1
'End Function
'
'Private Sub AddItem(Key As Variant, Item As Variant)
'    Dim Index As Long
'    Index = IndexOfKey(Key)
'
'    If Index >= 0 Then
'        ReplaceItemInCollection pItems, Item, Index
'    Else
'        pKeys.Add CStr(Key)
'        pItems.Add Item
'    End If
'End Sub
'
'Private Sub ReplaceItemInCollection(ByRef Coll As Collection, Item As Variant, Index As Long)
'    If Index >= 1 And Index <= Coll.Count Then
'        Coll.Remove Index
'
'        If Coll.Count = 0 Then
'            Coll.Add Item
'        ElseIf Index > Coll.Count Then
'            Coll.Add Item, After:=Index - 1
'        Else
'            Coll.Add Item, Before:=Index
'        End If
'    End If
'End Sub
'
'Private Function CollectionToArray(Coll As Collection) As Variant
'    Dim Arr As Variant
'    Dim i As Long
'
'    ' Collection is 1-based / Variant is 0-based
'    If Coll.Count > 0 Then
'        ReDim Arr(Coll.Count - 1)
'        For i = 1 To Coll.Count
'            If IsObject(Coll(i)) Then
'                Set Arr(i - 1) = Coll(i)
'            Else
'                Arr(i - 1) = Coll(i)
'            End If
'        Next i
'    Else
'        Arr = Array()
'    End If
'
'    CollectionToArray = Arr
'End Function

Private Sub Class_Initialize()
    Set pKeyValues = New Collection
End Sub

Private Sub Class_Terminate()
    Set pKeyValues = Nothing
End Sub
