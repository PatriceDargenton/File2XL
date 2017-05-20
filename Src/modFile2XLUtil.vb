
' File modFile2XLUtil.vb : Utility module for File2XL
' ----------------------

Module modFile2XLUtil

Public Function iNbOccurrences%(sTxt$, sOcc$)

    ' Return the number of items searched

    Dim iTxtLen% = sTxt.Length
    Dim iOccLen% = sOcc.Length
    Dim iPosMax% = iTxtLen - iOccLen
    Dim iNbOcc%, iPosNew%, iPos%
    While iPos <= iPosMax
        iPosNew = sTxt.IndexOf(sOcc, iPos, StringComparison.Ordinal) ' Ordinal : Exact (binary)
        If iPosNew = -1 Then Exit While
        iNbOcc += 1
        iPos = iPosNew + iOccLen
    End While
    'Debug.WriteLine("Nb " & sOcc & " = " & iNbOcc & " in " & sTxt)
    Return iNbOcc

End Function

End Module
