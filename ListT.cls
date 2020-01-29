VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.ListT")
Option Explicit

Private Type TListT
    ListT As Collection
    OptionStrict As Boolean
    ItemTypeName As String
End Type

Private this As TListT

Private Sub Class_Initialize()
    Set this.ListT = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.ListT = Nothing
End Sub

Public Property Let OptionStrict(ByVal value As Boolean)
    this.OptionStrict = value
End Property
Public Property Get OptionStrict() As Boolean
    OptionStrict = this.OptionStrict
End Property

Public Property Get Count() As Long
    Count = this.ListT.Count
End Property

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
'Gets the element at the specified index.

    If IsObjectType Then
        Set Item = this.ListT(index)
    Else
        Item = this.ListT(index)
    End If

End Property

Public Property Let Item(ByVal index As Long, ByVal value As Variant)
'Sets the element at the specified index.
    
    If Not ValidateItemType(value) Then RaiseErrorUnsafeType "Item(Let)", VBA.TypeName(value)

    RemoveAt index
    If index = Count Then
        Add value
    Else
        InsertAt index, value
    End If

End Property

Public Property Set Item(ByVal index As Long, ByVal value As Variant)
'Sets the element at the specified index.
    If Not ValidateItemType(value) Then RaiseErrorUnsafeType "Item(Set)", VBA.TypeName(value)

    RemoveAt index
    If index = Count Then
        Add value
    Else
        InsertAt index, value
    End If

End Property


Public Sub Add(ByVal value As Variant)
    If Not ValidateItemType(value) Then RaiseErrorUnsafeType "Add()", VBA.TypeName(value)
    this.ListT.Add value
End Sub

Public Function Contains(value As Variant) As Boolean
'Determines whether an element is in the List.

    Contains = (IndexOf(value) <> -1)

End Function

Public Function IndexOf(value As Variant, Optional Comparer As IComparator) As Long
'Searches for the specified object and returns the 1-based index of the first occurrence within the entire ListT

    If Count = 0 Then IndexOf = -1: Exit Function

    Dim isObjType As Boolean
    isObjType = IsObjectType
        
    Dim i As Long
        Dim found As Boolean
    For i = 1 To Count

        If isObjType Then

            found = (Comparer.CompareTo(value, Item(i)) = 0)

        Else

            found = (value = Item(i))

        End If

        If found Then IndexOf = i: Exit Function

    Next

    IndexOf = -1

End Function

Public Sub InsertBefore(ByVal beforeIndex As Long, ByVal value As Variant)

    If Not ValidateItemType(value) Then RaiseErrorUnsafeType "InsertBefore()", VBA.TypeName(value)
    
    Me.InsertAt beforeIndex, value
    
    'this.ListT.Add value, Before:=beforeIndex
    
End Sub

Public Sub InsertAfter(ByVal afterIndex As Long, ByVal value As Variant)
    
    If Not ValidateItemType(value) Then RaiseErrorUnsafeType "InsertAfter()", VBA.TypeName(value)
    
    Me.InsertAt afterIndex + 1, value
    
    'this.ListT.Add value, After:=afterIndex
    
End Sub

Public Sub InsertAt(ByVal index As Long, value As Variant)
'Inserts an element into the ListT at the specified index.

    Dim tmp As ListT
    Set tmp = GetRange(index, Count)

    RemoveRange index, Count

    Add value
    AddRange tmp

End Sub


Public Sub InsertArray(ByVal index As Long, values() As Variant)
'Inserts the specified elements into the ListT at the specified index.

    Dim tmp As ListT
    Set tmp = GetRange(index, Count)

    RemoveRange index, Count

    AddArray values
    AddRange tmp

End Sub


Public Sub InsertRange(ByVal index As Long, values As ListT)
'Inserts the specified elements into the ListT at the specified index.

    Dim tmp As ListT
    Set tmp = GetRange(index, Count)

    RemoveRange index, Count

    AddRange values
    AddRange tmp

End Sub


Public Sub InsertValues(ByVal index As Long, ParamArray values())
'Inserts the specified elements into the ListT at the specified index.

    Dim valuesArray() As Variant
    valuesArray = values

    InsertArray index, valuesArray

End Sub


Public Sub AddArray(values() As Variant)
'Adds the specified elements to the end of the List.

    Dim value As Variant, i As Long
    For i = LBound(values) To UBound(values)
        If ValidateItemType(value) Then
            this.ListT.Add values(i)
        Else
            RaiseErrorUnsafeType "AddArray()", VBA.TypeName(value)
        End If
    Next

End Sub


Public Sub AddRange(ByRef values As ListT)
'Adds the specified elements to the end of the ListT.

    Dim value As Variant
    For Each value In values
        Add value
    Next

End Sub

Public Sub AddValues(ParamArray values())
'Adds the specified elements to the end of the ListT.

    Dim value As Variant, i As Long
    For i = LBound(values) To UBound(values)
        Add values(i)
    Next

End Sub

Public Sub RemoveAt(ByVal index As Long)
    this.ListT.Remove index
End Sub

'Public Sub RemoveAt(ByVal Index As Long)
'
'    Dim tmp As ListT
'    Set tmp = GetRange(Index + 1, Count)
'
'    RemoveRange Index, Count
'    AddRange tmp
'
'    'this.ListT.Remove Index
'End Sub


Public Sub RemoveRange(ByVal startIndex As Long, ByVal valuesCount As Long)
'Removes a range of elements from the UniqueList.
    'If valuesCount > this.UniqueList.Count Then Err.Raise
    
    Dim i As Long
    On Error GoTo CleanFail
    For i = valuesCount To startIndex Step -1
        this.ListT.Remove i
    Next
    
'    For i = startIndex To (startIndex + valuesCount - 1)
'        this.ListT.Remove i
'    Next

CleanExit:
    Exit Sub
    
CleanFail:
    Stop
    Debug.Print Err.Number
    Debug.Print Err.Description
    Resume CleanExit
    
End Sub


Public Sub Clear()
    Set this.ListT = Nothing
    Set this.ListT = New Collection
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.ListT.[_NewEnum]
End Property


Public Function GetRange(ByVal index As Long, ByVal valuesCount As Long) As ListT
'Creates a copy of a range of elements in the source ListT.

    Dim result As ListT
    If index > Count Then Err.Raise 9 'index out of range

    Dim lastIndex As Long
    lastIndex = IIf(index + valuesCount > Count, Count, index + valuesCount)

    Set result = New ListT

    Dim i As Long
    For i = index To lastIndex
        result.Add Item(i)
    Next

    Set GetRange = result

End Function

Public Sub Reverse()
'Reverses the Priority of the elements in the entire List.

    Dim tmp As New ListT
    Do Until Count = 0

        tmp.Add Item(Count)
        RemoveAt Count

    Loop

    AddRange tmp

End Sub


Public Sub Sort(Comparer As IComparator, Optional ByVal ascending As Boolean = True)
    
    If IsObjectType Then
        If ascending Then
            QuickSortObjectsAsc Comparer, 1, this.ListT.Count
        Else
            QuickSortObjectsDesc Comparer, 1, this.ListT.Count
        End If
    Else
        If ascending Then
            QuickSortValuesAsc Comparer, 1, this.ListT.Count
        Else
            QuickSortValuesDesc Comparer, 1, this.ListT.Count
        End If
    End If
    
    
End Sub



Private Sub QuickSortObjectsAsc(Comparer As IComparator, ByVal Min As Long, ByVal Max As Long)
'Attribute NewEnum.VB_MemberFlags = "40"
    Dim minVar As Variant, maxVar As Variant
    Dim minComparable As Object, maxComparable As Object

    Dim i As Long, j As Long
    i = Min
    j = Max

    Dim pivot As Variant
    Set pivot = this.ListT((Min + Max) \ 2)

    Do While i <= j

        Set minComparable = this.ListT.Item(i)

        Do While Comparer.CompareTo(minComparable, pivot) < 0 And i < Max
            i = i + 1

            Set minComparable = this.ListT.Item(i)

        Loop

        Set maxComparable = this.ListT.Item(j)

        Do While Comparer.CompareTo(maxComparable, pivot) > 0 And j > Min
            j = j - 1

            Set maxComparable = this.ListT.Item(j)

        Loop

        If i <= j Then

            Set minVar = this.ListT(i)
            Set maxVar = this.ListT(j)
            
            this.ListT.Remove i
            If i <= this.ListT.Count Then this.ListT.Add maxVar, Before:=i Else this.ListT.Add maxVar

            this.ListT.Remove j
            If j <= this.ListT.Count Then this.ListT.Add minVar, Before:=j Else this.ListT.Add minVar

            ' Move to next positions
            i = i + 1
            j = j - 1

        End If

    Loop

    If Min < j Then QuickSortObjectsAsc Comparer, Min, j
    If i < Max Then QuickSortObjectsAsc Comparer, i, Max

End Sub


Private Sub QuickSortObjectsDesc(Comparer As IComparator, ByVal Min As Long, ByVal Max As Long)
'Attribute NewEnum.VB_MemberFlags = "40"

    Dim minVar As Variant, maxVar As Variant
    Dim minComparable As Object, maxComparable As Object
    Dim minKey As String, maxKey As String

    Dim i As Long, j As Long
    i = Min
    j = Max
    
    Dim pivot As Variant
    Set pivot = this.ListT((Min + Max) \ 2)
    
    Do While i <= j

        Set minComparable = this.ListT.Item(i)
        Do While Comparer.CompareTo(minComparable, pivot) > 0 And i < Max
            i = i + 1
            
            Set minComparable = this.ListT.Item(i)

        Loop
            
        Set maxComparable = this.ListT.Item(j)
        Do While Comparer.CompareTo(maxComparable, pivot) < 0 And j > Min
            j = j - 1

            Set maxComparable = this.ListT.Item(j)

        Loop

        If i <= j Then

            Set maxVar = this.ListT(j)
            Set minVar = this.ListT(i)
            
            this.ListT.Remove i
            If i <= this.ListT.Count Then this.ListT.Add maxVar, Before:=i Else this.ListT.Add maxVar

            this.ListT.Remove j
            If j <= this.ListT.Count Then this.ListT.Add minVar, Before:=j Else this.ListT.Add minVar

            ' Move to next positions
            i = i + 1
            j = j - 1

        End If

    Loop

    If Min < j Then QuickSortObjectsDesc Comparer, Min, j
    If i < Max Then QuickSortObjectsDesc Comparer, i, Max

End Sub


Private Sub QuickSortValuesAsc(Comparer As IComparator, ByVal Min As Long, ByVal Max As Long)
'Attribute NewEnum.VB_MemberFlags = "40"
    Dim minVar As Variant, maxVar As Variant
    Dim minComparable As Variant, maxComparable As Variant
    Dim minKey As String, maxKey As String

    Dim i As Long, j As Long
    i = Min
    j = Max
    
    Dim pivot As Variant
    pivot = this.ListT((Min + Max) \ 2)
    
    Do While i <= j

        minComparable = this.ListT.Item(i)
        Do While Comparer.CompareTo(minComparable, pivot) < 0 And i < Max
            i = i + 1
            
            minComparable = this.ListT.Item(i)

        Loop
            
        maxComparable = this.ListT.Item(j)
        Do While Comparer.CompareTo(maxComparable, pivot) > 0 And j > Min
            j = j - 1

            maxComparable = this.ListT.Item(j)

        Loop

        If i <= j Then

            maxVar = this.ListT(j)
            minVar = this.ListT(i)
            
            this.ListT.Remove i
            If i <= this.ListT.Count Then this.ListT.Add maxVar, Before:=i Else this.ListT.Add maxVar

            this.ListT.Remove j
            If j <= this.ListT.Count Then this.ListT.Add minVar, Before:=j Else this.ListT.Add minVar

            'Move to next positions
            i = i + 1
            j = j - 1

        End If

    Loop

    If Min < j Then QuickSortValuesAsc Comparer, Min, j
    If i < Max Then QuickSortValuesAsc Comparer, i, Max

End Sub


Private Sub QuickSortValuesDesc(ByRef Comparer As IComparator, ByVal Min As Long, ByVal Max As Long)
'Attribute NewEnum.VB_MemberFlags = "40"
    Dim minVar As Variant, maxVar As Variant
    Dim minComparable As Variant, maxComparable As Variant
    
    If Min = -1 Then Min = 1
    If Max = -1 Then Max = this.ListT.Count

    Dim i As Long, j As Long
    i = Min
    j = Max
    
    Dim pivot As Variant
    pivot = this.ListT((Min + Max) \ 2)
    
    Do While i <= j

        minComparable = this.ListT.Item(i)
        Do While Comparer.CompareTo(minComparable, pivot) > 0 And i < Max
            i = i + 1
            
            minComparable = this.ListT.Item(i)

        Loop
            
        maxComparable = this.ListT.Item(j)
        Do While Comparer.CompareTo(maxComparable, pivot) < 0 And j > Min
            j = j - 1

            maxComparable = this.ListT.Item(j)

        Loop

        If i <= j Then

            maxVar = this.ListT(j)
            minVar = this.ListT(i)

            this.ListT.Remove i
            If i <= this.ListT.Count Then this.ListT.Add maxVar, Before:=i Else this.ListT.Add maxVar

            this.ListT.Remove j
            If j <= this.ListT.Count Then this.ListT.Add minVar, Before:=j Else this.ListT.Add minVar

            i = i + 1
            j = j - 1

        End If

    Loop

    If Min < j Then QuickSortValuesDesc Comparer, Min, j
    If i < Max Then QuickSortValuesDesc Comparer, i, Max

End Sub

Private Function IsObjectType() As Boolean
'Attribute NewEnum.VB_MemberFlags = "40"
    If this.ListT.Count = 0 Then IsObjectType = False: Exit Function
    IsObjectType = IsObject(this.ListT(1))
End Function


Private Function ValidateItemType(value As Variant) As Boolean

   If this.ItemTypeName = vbNullString Then this.ItemTypeName = VBA.TypeName(value)
   ValidateItemType = IsTypeSafe(value)

End Function

Private Function IsTypeSafe(value As Variant) As Boolean

    IsTypeSafe = this.ItemTypeName = vbNullString Or this.ItemTypeName = VBA.TypeName(value)
    If IsTypeSafe Or this.OptionStrict Then Exit Function
    
        On Error GoTo CleanFail
        Select Case VarType(value)

                Case vbInteger
                        value = CInt(value)
                        IsTypeSafe = True
                Case vbLong
                        value = CLng(value)
                        IsTypeSafe = True
                        
                Case vbSingle
                        value = CSng(value)
                        IsTypeSafe = True
                        
                Case vbDouble
                        value = CDbl(value)
                        IsTypeSafe = True
                        
                Case vbCurrency
                        value = CCur(value)
                        IsTypeSafe = True
                        
                Case vbDate
                        value = CDate(value)
                        IsTypeSafe = True
                        
                Case vbString
                        value = CInt(value)
                        IsTypeSafe = True
                        
                Case vbObject
                        value = value
                        IsTypeSafe = True
                        
                Case vbBoolean
                        value = CBool(value)
                        IsTypeSafe = True
                        
                Case vbDecimal
                        value = CDec(value)
                        IsTypeSafe = True
                        
                Case vbByte
                        value = CByte(value)
                        IsTypeSafe = True
                
                Case Else
                        IsTypeSafe = False
                        
        End Select
        
CleanExit:
        Exit Function
        
CleanFail:
    'swallow overflow errors:
    If Err.Number = 6 Then Err.Clear
    IsTypeSafe = False
    Resume CleanExit

End Function
        
Private Sub RaiseErrorUnsafeType(member As String, suppliedType As String)
    Err.Raise 13, suppliedType, "Type Mismatch. Expected: " & member & " " & suppliedType & " was supplied."
End Sub

