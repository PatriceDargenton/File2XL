
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

    If String.IsNullOrEmpty(sort) Then Exit Sub

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
        If memberName.EndsWith(" desc", StringComparison.OrdinalIgnoreCase) Then
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

Public Function Compare(x As Object, y As Object) As Integer _
    Implements IComparer.Compare
    ' Implementation of IComparer.Compare
    Return Compare(CType(x, T), CType(y, T))
End Function

Public Function Compare(x As T, y As T) As Integer _
    Implements IComparer(Of T).Compare

    ' Implementation of IComparer(Of T).Compare

    ' Deal with the simplest cases first.
    If x Is Nothing Then
        ' Two null objects are equal.
        If y Is Nothing Then Return 0
        ' A null object is less than any non-null object.
        Return -1
    ElseIf y Is Nothing Then
        ' Any non-null object is greater than a null object.
        Return 1
    End If

    ' Iterate over all the sort keys.
    For i As Integer = 0 To sortKeys.Length - 1
        Dim oValue_x As Object, oValue_y As Object
        Dim sortKey As SortKey = sortKeys(i)
        ' Read either the field or the property.
        If sortKey.FieldInfo IsNot Nothing Then
            oValue_x = sortKey.FieldInfo.GetValue(x)
            oValue_y = sortKey.FieldInfo.GetValue(y)
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
            oValue_x = sortKey.PropertyInfo.GetValue(x, Nothing)
            oValue_y = sortKey.PropertyInfo.GetValue(y, Nothing)
        End If

        Dim iRes%
        If oValue_x Is Nothing And oValue_y Is Nothing Then
            ' Two null objects are equal.
            iRes = 0
        ElseIf oValue_x Is Nothing Then
            ' A null object is always less than a non-null object.
            iRes = -1
        ElseIf oValue_y Is Nothing Then
            ' Any object is greater than a null object.
            iRes = 1
        Else
            ' Compare the two values, assuming that they support IComparable.
            iRes = DirectCast(oValue_x, IComparable).CompareTo(oValue_y)
        End If

        ' If values are different, return this value to caller.
        If iRes <> 0 Then
            ' Negate it if sort direction is descending.
            If sortKey.Descending Then iRes = -iRes
            Return iRes
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