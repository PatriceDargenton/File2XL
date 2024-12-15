
' File clsSortDic.vb : Sortable dictionary class
' ------------------

Imports System.Runtime.Serialization

<Serializable> _
Public Class SortDic(Of TKey, TValue) : Inherits Dictionary(Of TKey, TValue)

    Sub New()
    End Sub

    ' SYSLIB0051: remove this constructor: obsolete
    'Protected Sub New(info As SerializationInfo, context As StreamingContext)
    '    MyBase.New(info, context)
    'End Sub

    Public Function Sort(Optional sSorting$ = "") As TValue()

        ' Sort the dictionary and return sorted elements

        Dim iNbLines% = Me.Count
        Dim arrayTvalue(iNbLines - 1) As TValue
        Dim iNumLine% = 0
        For Each kvp As KeyValuePair(Of TKey, TValue) In Me
            arrayTvalue(iNumLine) = kvp.Value
            iNumLine += 1
        Next

        ' If no sorting is specified, simply return the array
        If String.IsNullOrEmpty(sSorting) Then Return arrayTvalue

        ' Sort the dictionary
        Dim comp As New UniversalComparer(Of TValue)(sSorting)
        Array.Sort(Of TValue)(arrayTvalue, comp)
        Return arrayTvalue

    End Function

End Class