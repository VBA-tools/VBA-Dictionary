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
' Dictionary v1.0.2
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

Private pKeys As Collection
Private pItems As Collection
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
        Err.Raise 5
    End If
    
    pCompareMode = Value
End Property

Public Property Get Count() As Long
    Count = pKeys.Count
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Index As Long
    Index = IndexOfKey(Key)
    
    If Index >= 0 Then
        If IsObject(pItems(Index)) Then
            Set Item = pItems(Index)
        Else
            Item = pItems(Index)
        End If
    Else
        ' Not found -> Returns Empty
    End If
End Property
Public Property Let Item(Key As Variant, Value As Variant)
    AddItem Key, Value
End Property
Public Property Set Item(Key As Variant, Value As Variant)
    AddItem Key, Value
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    Dim Index As Long
    Index = IndexOfKey(Previous)
    
    If Index >= 0 Then
        ReplaceItemInCollection pKeys, CStr(Updated), Index
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {String} Key
' @param {Item} Variant
' --------------------------------------------- '
Public Sub Add(Key As Variant, Item As Variant)
    If Not Me.Exists(Key) Then
        pKeys.Add CStr(Key)
        pItems.Add Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
End Sub

''
' Check if an item exists for the given key
'
' @param {String} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
    Exists = IndexOfKey(Key) >= 0
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
    Items = CollectionToArray(pItems)
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
    Keys = CollectionToArray(pKeys)
End Function

''
' Remove an item for the given key
'
' @param {String} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
    Dim Index As Long
    Index = IndexOfKey(Key)
    
    If Index > -1 Then
        pKeys.Remove Index
        pItems.Remove Index
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
    Set pKeys = New Collection
    Set pItems = New Collection
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function IndexOfKey(Key As Variant) As Long
    Dim i As Long
    Key = CStr(Key)
    
    For i = 1 To pKeys.Count
        If VBA.StrComp(pKeys(i), Key, Me.CompareMode) = 0 Then
            IndexOfKey = i
            Exit Function
        End If
    Next i
    
    ' Not found
    IndexOfKey = -1
End Function

Private Sub AddItem(Key As Variant, Item As Variant)
    Dim Index As Long
    Index = IndexOfKey(Key)
    
    If Index >= 0 Then
        ReplaceItemInCollection pItems, Item, Index
    Else
        pKeys.Add CStr(Key)
        pItems.Add Item
    End If
End Sub

Private Sub ReplaceItemInCollection(ByRef Coll As Collection, Item As Variant, Index As Long)
    If Index >= 1 And Index <= Coll.Count Then
        Coll.Remove Index
        
        If Coll.Count = 0 Then
            Coll.Add Item
        ElseIf Index > Coll.Count Then
            Coll.Add Item, After:=Index - 1
        Else
            Coll.Add Item, Before:=Index
        End If
    End If
End Sub

Private Function CollectionToArray(Coll As Collection) As Variant
    Dim Arr As Variant
    Dim i As Long
    
    ' Collection is 1-based / Variant is 0-based
    If Coll.Count > 0 Then
        ReDim Arr(Coll.Count - 1)
        For i = 1 To Coll.Count
            If IsObject(Coll(i)) Then
                Set Arr(i - 1) = Coll(i)
            Else
                Arr(i - 1) = Coll(i)
            End If
        Next i
    Else
        Arr = Array()
    End If
    
    CollectionToArray = Arr
End Function

Private Sub Class_Initialize()
    Set pKeys = New Collection
    Set pItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set pKeys = Nothing
    Set pItems = Nothing
End Sub
