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

' KeyValue 0: Key, 1: Value
Private pKeyValues As Collection
Private pKeys() As Variant
Private pItems() As Variant
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
    Items = pItems
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
    Keys = pKeys
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
    Erase pKeys
    Erase pItems
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function GetKeyValue(Key As Variant) As Variant
    On Error Resume Next
    GetKeyValue = pKeyValues(GetFormattedKey(Key))
End Function

Private Sub AddKeyValue(Key As Variant, Value As Variant, Optional Index As Long = -1)
    If Me.Count = 0 Then
        ReDim Preserve pKeys(0 To 0)
        ReDim Preserve pItems(0 To 0)
    Else
        ReDim Preserve pKeys(0 To UBound(pKeys) + 1)
        ReDim Preserve pItems(0 To UBound(pItems) + 1)
    End If
    
    Dim FormattedKey As String
    FormattedKey = GetFormattedKey(Key)
    
    If Index > 0 And Index <= pKeyValues.Count Then
        Dim i As Long
        For i = UBound(pKeys) To Index Step -1
            pKeys(i) = pKeys(i - 1)
            If IsObject(pItems(i - 1)) Then
                Set pItems(i) = pItems(i - 1)
            Else
                pItems(i) = pItems(i - 1)
            End If
        Next i
        
        pKeys(Index - 1) = Key
        If IsObject(Value) Then
            Set pItems(Index - 1) = Value
        Else
            pItems(Index - 1) = Value
        End If
        
        pKeyValues.Add Array(FormattedKey, Key, Value), FormattedKey, Before:=Index
    Else
        pKeys(UBound(pKeys)) = Key
        If IsObject(Value) Then
            Set pItems(UBound(pItems)) = Value
        Else
            pItems(UBound(pItems)) = Value
        End If
        
        pKeyValues.Add Array(FormattedKey, Key, Value), FormattedKey
    End If
End Sub

Private Sub ReplaceKeyValue(KeyValue As Variant, Key As Variant, Value As Variant)
    Dim Index As Long
    Dim i As Integer
    
    For i = 0 To UBound(pKeys)
        If pKeys(i) = KeyValue(1) Then
            Index = i + 1
            Exit For
        End If
    Next i
    
    ' Remove existing value
    RemoveKeyValue KeyValue, Index
    
    ' Add new key value back
    AddKeyValue Key, Value, Index
End Sub

Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal Index As Long = -1)
    Dim i As Long
    If Index = -1 Then
        For i = 0 To UBound(pKeys)
            If pKeys(i) = KeyValue(1) Then
                Index = i
            End If
        Next i
    Else
        Index = Index - 1
    End If
    
    If Index >= 0 And Index <= UBound(pKeys) Then
        For i = Index To UBound(pKeys) - 1
            pKeys(i) = pKeys(i + 1)
            
            If IsObject(pItems(i + 1)) Then
                Set pItems(i) = pItems(i + 1)
            Else
                pItems(i) = pItems(i + 1)
            End If
        Next i
        
        If UBound(pKeys) = 0 Then
            Erase pKeys
            Erase pItems
        Else
            ReDim Preserve pKeys(0 To UBound(pKeys) - 1)
            ReDim Preserve pItems(0 To UBound(pItems) - 1)
        End If
    End If
    
    pKeyValues.Remove KeyValue(0)
End Sub

Private Function GetFormattedKey(Key As Variant) As String
    GetFormattedKey = CStr(Key)
    If Me.CompareMode = CompareMethod.BinaryCompare Then
        ' Collection does not have method of setting key comparison
        ' So case-sensitive keys aren't supported by default
        ' -> Approach: Append lowercase characters to original key
        '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
        '    Won't work in very strange cases, but should work for now
        '    AbBb -> AbBb__bb matches AbbB -> AbbB__bb
        Dim Lowercase As String
        Lowercase = ""
        
        Dim i As Integer
        Dim Ascii As Integer
        Dim Char As String
        For i = 1 To Len(GetFormattedKey)
            Char = VBA.Mid$(GetFormattedKey, i, 1)
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

Private Sub Class_Initialize()
    Set pKeyValues = New Collection
    Erase pKeys
    Erase pItems
End Sub

Private Sub Class_Terminate()
    Set pKeyValues = Nothing
End Sub
