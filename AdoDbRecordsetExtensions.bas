Attribute VB_Name = "AdoDbRecordsetExtensions"
'@Folder("Framework.DataAccess.Common.Extensions")
Option Explicit

Public Function RecordsetToDictionary(ByRef recordset As ADODB.recordset, _
                                      ByVal keyFieldName As String) As Object

    On Error GoTo CleanFail
    Dim dictOut As Object
    Set dictOut = CreateObject("Scripting.Dictionary")
    
    Dim arryTemp() As Variant
    ReDim arryTemp(recordset.Fields.Count - 1)
    
    Dim varKey As Variant
    Dim i As Long
    Dim fld As ADODB.Field

    Do While Not recordset.EOF
    
        varKey = recordset.Fields(keyFieldName).value
        
        If Not dictOut.Exists(varKey) Then
            
            For Each fld In recordset.Fields
                
                arryTemp(i) = fld.value
                i = i + 1
            Next
            
            dictOut(varKey) = arryTemp
            i = Empty
            
        End If
        
        recordset.MoveNext
    Loop

    If Not ((recordset.cursorType And adOpenForwardOnly) = adOpenForwardOnly) Then recordset.MoveFirst
    
    Set RecordsetToDictionary = dictOut

CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit
    
End Function


Public Function RecordsetToArray(ByRef recordset As ADODB.recordset, ByVal transpose As Boolean) As Variant
    
    On Error GoTo CleanFail
    If transpose Then
        Dim arryTemp() As Variant
        ReDim arryTemp(recordset.RecordCount - 1, recordset.Fields.Count - 1)

        Dim arrayRecords As Variant
        arrayRecords = recordset.GetRows()
        
        Dim j As Long, i As Long
        For j = 0 To UBound(arrayRecords, 2)
            For i = 0 To UBound(arrayRecords, 1)
                arryTemp(j, i) = arrayRecords(i, j)
            Next i
        Next j
        
        RecordsetToArray = arryTemp
    
    Else
        RecordsetToArray = recordset.GetRows()
    
    End If
    
CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit

End Function

Public Function FieldNamesToArray(ByRef recordset As ADODB.recordset) As Variant


    On Error GoTo CleanFail
    With recordset
        Dim arryFieldNames() As Variant
        ReDim arryFieldNames(.Fields.Count)
        
        Dim i As Long
        For i = 0 To .Fields.Count - 1
            arryFieldNames(i) = .Fields.Item(i).name
        Next i

    End With
    
    FieldNamesToArray = arryFieldNames
    
CleanExit:
    Exit Function

CleanFail:
    Resume CleanExit
    
End Function


Public Sub FieldNamesToRange(ByRef recordset As ADODB.recordset, ByRef rangeOut As Range)

    Dim arryFieldNames() As Variant
    arryFieldNames() = FieldNamesToArray(recordset)
    
    rangeOut.Resize(1, UBound(arryFieldNames) + 1).Value2 = Application.transpose(Application.transpose(arryFieldNames))
    
End Sub


