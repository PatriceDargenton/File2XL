
' File UniversalComparer.vb : Generic comparer for any class
' -------------------------

Imports System.Collections.Generic
Imports System.Reflection

' http://archive.visualstudiomagazine.com/2005_02/magazine/columns/net2themax/Listing2.aspx

Public Class UniversalComparer(Of T) : Implements IComparer, IComparer(Of T)

Private sortKeys() As SortKey
Private m_bMsg As Boolean = False
Private m_sSorting$ = ""

Public Sub New(sort As String)

    m_sSorting = sort

    Dim type As Type = GetType(T)
    ' Split the list of properties.
    Dim props() As String = sort.Split(","c)
    ' Prepare the array that holds information on sort criteria.
    ReDim sortKeys(props.Length - 1)

    ' Parse the sort string.
    For i As Integer = 0 To props.Length - 1
        ' Get the N-th member name.
        Dim memberName As String = props(i).Trim()
        If memberName.ToLower().EndsWith(" desc") Then
            ' Discard the DESC qualifier.
            sortKeys(i).Descending = True
            memberName = memberName.Remove(memberName.Length - 5).TrimEnd()
        End If
        ' Search for a field or a property with this name.
        sortKeys(i).FieldInfo = type.GetField(memberName)
        sortKeys(i).sMemberName = memberName
        If sortKeys(i).FieldInfo Is Nothing Then
            sortKeys(i).PropertyInfo = type.GetProperty(memberName)
        End If
    Next i

End Sub

Public Function Compare(o1 As Object, o2 As Object) As Integer _
    Implements IComparer.Compare
    ' Implementation of IComparer.Compare
    Return Compare(CType(o1, T), CType(o2, T))
End Function

Public Function Compare(o1 As T, o2 As T) As Integer _
    Implements IComparer(Of T).Compare

    ' Implementation of IComparer(Of T).Compare

    ' Deal with the simplest cases first.
    If o1 Is Nothing Then
        ' Two null objects are equal.
        If o2 Is Nothing Then Return 0
        ' A null object is less than any non-null object.
        Return -1
    ElseIf o2 Is Nothing Then
        ' Any non-null object is greater than a null object.
        Return 1
    End If

    ' Iterate over all the sort keys.
    For i As Integer = 0 To sortKeys.Length - 1
        Dim value1 As Object, value2 As Object
        Dim sortKey As SortKey = sortKeys(i)
        ' Read either the field or the property.
        If sortKey.FieldInfo IsNot Nothing Then
            value1 = sortKey.FieldInfo.GetValue(o1)
            value2 = sortKey.FieldInfo.GetValue(o2)
        Else
            If IsNothing(sortKey.PropertyInfo) Then
                If Not m_bMsg Then
                    MsgBox( _
                        "A sorting key can't be found : the specified field doesn't exists" & vbLf & _
                        "or is not in public scope !" & vbLf & _
                        GetType(T).ToString & " : " & sortKeys(i).sMemberName & " : " & m_sSorting, _
                        MsgBoxStyle.Critical, "UniversalComparer:Compare")
                    m_bMsg = True
                End If
                Return 0
            End If
            value1 = sortKey.PropertyInfo.GetValue(o1, Nothing)
            value2 = sortKey.PropertyInfo.GetValue(o2, Nothing)
        End If

        Dim res As Integer
        If value1 Is Nothing And value2 Is Nothing Then
            ' Two null objects are equal.
            res = 0
        ElseIf value1 Is Nothing Then
            ' A null object is always less than a non-null object.
            res = -1
        ElseIf value2 Is Nothing Then
            ' Any object is greater than a null object.
            res = 1
        Else
            ' Compare the two values, assuming that they support IComparable.
            res = DirectCast(value1, IComparable).CompareTo(value2)
        End If

        ' If values are different, return this value to caller.
        If res <> 0 Then
            ' Negate it if sort direction is descending.
            If sortKey.Descending Then res = -res
            Return res
        End If
    Next i

    ' If we get here the two objects are equal.
    Return 0

End Function

Private Structure SortKey ' Nested type to store detail on sort keys
    Public FieldInfo As FieldInfo
    Public PropertyInfo As PropertyInfo
    ' True if sort is descending.
    Public Descending As Boolean
    Public sMemberName$
End Structure

End Class