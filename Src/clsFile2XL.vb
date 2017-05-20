
' File clsFile2XL.vb : Open a csv file into MS-Excel with pre-formatted cells
' ------------------

Imports System.Text ' For StringBuilder
Imports NPOI.XSSF.UserModel ' For XSSFWorkbook, XSSFSheet : Excel 2007
Imports NPOI.HSSF.UserModel ' For HSSFWorkbook, HSSFSheet : Excel 2003
Imports NPOI.HSSF.Model
Imports NPOI.SS.Util ' For CellRangeAddress
Imports NPOI.SS.UserModel ' For FillPattern
Imports NPOI.HSSF.Util ' For HSSFColor
Imports System.Runtime.CompilerServices ' For MethodImpl(MethodImplOptions.AggressiveInlining)

Public Class clsPrm

    Public sFieldDelimiters$, sDefaultDelimiter$
    Public bUseXls As Boolean
    Public bUseXlsx As Boolean
    Public bAlertForNoDelimiterFound As Boolean = True
    Public bUseQuotesCommaQuotesDelimiter As Boolean = True
    Public bMsgBoxAlert As Boolean = True
    Public iNbFrozenColumns%
    Public iNbLinesAnalyzed%
    Public bCreateStandardSheet As Boolean
    Public bPreferMultipleDelimiter As Boolean ' For example, prefer "," to ,
    Public bAutosizeColumns As Boolean
    Public iMinColumnWidth% ' After autozise 20/05/2017
    Public iMaxColumnWidth% ' After autozise 20/05/2017
    Public bRemoveNULL As Boolean ' Replace PhpMyAdmin NULL by empty 28/04/2017

End Class

Public Class clsFile2XL

Public m_bXlsx As Boolean = False
Public m_sDestPathXls$, m_sDestPathXlsx$
Public m_bOnlyTextFields As Boolean = True ' Check if there are only text fields or not, and store them here

Public Const iNbCarMaxCell% = 32767
Public Const iNbLinesMaxExcel2003% = 65536
Public Const iNbLinesMaxExcel2007% = 1048576
Const iNbColMaxExcel2003% = 256
Const iNbColMaxExcel2007% = 16384
'Dim m_iNbColMaxExcel% = iNbColMaxExcel2003
Const iNbColMaxAutoFilterExcel2003NPOI% = 255 ' Bug NPOI : il should be 256

Const sMsgNextColumnsIgnored$ = "(File2XL: Next columns have been ignored)"
Const sMsgNextLinesIgnored$ = "(File2XL: Next lines have been ignored)"
Dim sMsgNextCarIgnored$ = "(File2XL: " & iNbCarMaxCell & _
    " characters reached, next characters have been ignored) "
Dim sMsgMaxCarCell$ = "The number of characters in a cell exceeds" & vbLf & _
    " the maximum allowed (" & iNbCarMaxCell & ")." & vbLf & _
    "Next characters will be ignored."
Dim m_sMsgMaxColumns$ = ""
Dim m_sMsgMaxLines$ = ""

Private WithEvents m_lineDeleg As New clsDelegLine
Private m_lines As List(Of String)
Private m_wb As HSSFWorkbook, m_sh As HSSFSheet, m_shStdr As HSSFSheet ' Excel 2003
Private m_wbXlsx As XSSFWorkbook, m_shXlsx As XSSFSheet, m_shStdrXlsx As XSSFSheet ' Excel >= 2007

Private m_numericCellStyleXls, m_numericCellStyleXlsx As ICellStyle

Private m_iNumLine%
Private m_iNbColMaxFound% ' Not used
Private m_iNbFilledColMaxFound%
Private m_bAlertLineMax, m_bAlertColumnMax, m_bAlertCellTextLengthMax As Boolean
Private m_delegMsg As clsDelegMsg
Private m_prm As clsPrm

Private m_bDetectColumnType As Boolean
Private m_splitLines As List(Of List(Of String))
Private m_lstFields As List(Of clsField)

#Region "Classes"

Private Class clsFieldType
    Public Const sNumericC2P$ = "NumericC2P"
    Public Const sNumericP2C$ = "NumericP2C" ' Period to Comma
    Public Const sNumeric$ = "Numeric"
    Public Const sNumericWithQuotes$ = "NumericWithQuotes"
    Public Const sNumericC2PWithQuotes$ = "NumericC2PWithQuotes"
    Public Const sNumericP2CWithQuotes$ = "NumericP2CWithQuotes"
    Public Const sText$ = "Text"
    Public Const sTextWithQuotes$ = "TextWithQuotes"
End Class

Private Class clsField

    ' sField and iNumField are used only in debug mode
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")> _
    Public sField$, iNumField%
    Public sType$, iNbOcc%
    Public bCanEndWithMinus As Boolean = False ' Numeric followed by -
    Public Sub New(iNumField0%, sField0$, sType0$)
        iNumField = iNumField0
        sField = sField0
        sType = sType0
        iNbOcc = 1
    End Sub
End Class

Private Class clsOcc
    Public s$
    Public iNbOcc%, iOccLength%

    ' This field is used in sorting using a string, e.g.: "iWeight DESC, iNbOcc DESC, iOccLength DESC"
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>
    Public iWeight%

    Public Sub New(s0$, iNbOcc0%, bPreferMultipleDelimiter As Boolean)
        s = s0
        iNbOcc = iNbOcc0
        iOccLength = s.Length
        If bPreferMultipleDelimiter Then
            iWeight = iNbOcc * iOccLength ' Increase the weight as the delimiter length
        Else
            iWeight = iNbOcc
        End If
    End Sub
End Class

#End Region

Public Function bRead(prm As clsPrm, sPath$, delegMsg As clsDelegMsg) As Boolean

    If prm Is Nothing Then Return False
    If delegMsg Is Nothing Then Return False

    m_prm = prm
    If Not m_prm.bUseXls AndAlso Not m_prm.bUseXlsx Then
        If bDebug Then Stop
        m_prm.bUseXlsx = True
    End If

    m_delegMsg = delegMsg
    m_delegMsg.m_bIgnoreNextLines = False
    m_delegMsg.m_bCancel = False

    If Not bFileIsWritable(m_sDestPathXls, bNonExistentOk:=True, bPromptRetry:=True) Then Return False
    If Not bFileIsWritable(m_sDestPathXlsx, bNonExistentOk:=True, bPromptRetry:=True) Then Return False

    Dim encod = GetEncoding(sPath)
    ' If encoding is ASCII, set the Latin alphabet to preserve for example accents
    ' Default = System.Text.SBCSCodePageEncoding = Encoding.GetEncoding(1252)
    If encod Is Encoding.ASCII Then encod = Encoding.Default

    delegMsg.ShowMsg("Reading first lines...")
    m_lines = New List(Of String)
    Dim bHeader As Boolean
    Dim iNbLines% = 0
    Dim iNbColumns% = 0
    Dim sFieldDelimiter$ = String.Empty
    If Not bReadFileGeneric(sFieldDelimiter, bHeader, sPath, m_lineDeleg, delegMsg, _
        iNbLines, iNbColumns, _
        bOnlyFirstLines:=True, encod:=encod, iNbLinesAnalyzed:=prm.iNbLinesAnalyzed) Then Return False

    delegMsg.ShowLongMsg("")
    delegMsg.ShowMsg("Searching probable delimiter...")
    FindProbDelimiter(prm.sFieldDelimiters, prm.sDefaultDelimiter, sFieldDelimiter)

    ' Detecting column type
    m_splitLines = New List(Of List(Of String))
    m_bDetectColumnType = True
    If Not bReadFileGeneric(sFieldDelimiter, bHeader, sPath, m_lineDeleg, delegMsg, _
        iNbLines, iNbColumns, _
        bOnlyFirstSplitLines:=True, encod:=encod, iNbLinesAnalyzed:=prm.iNbLinesAnalyzed) Then Return False
    m_bDetectColumnType = False
    delegMsg.ShowMsg("Searching columns type...")
    FindColumnsType(m_lstFields, m_bOnlyTextFields)
    If m_bOnlyTextFields Then m_prm.bCreateStandardSheet = False

    delegMsg.ShowMsg("Initializing Excel library...")

    ' Read the file using the probable delimiter now

    m_bXlsx = False : If Not m_prm.bUseXls Then m_bXlsx = True
    UpdateMsg()

    If m_prm.bUseXls Then
        m_wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook()) ' Excel 2003
        m_sh = DirectCast(m_wb.CreateSheet(sTxtSheet), HSSFSheet)
        m_sh.CreateFreezePane(prm.iNbFrozenColumns, 1)
        If m_prm.bCreateStandardSheet Then
            m_shStdr = DirectCast(m_wb.CreateSheet(sStdrSheet), HSSFSheet)
            m_shStdr.CreateFreezePane(prm.iNbFrozenColumns, 1)
            SetNumericStyle(m_numericCellStyleXls, bExcel2007:=False)
        End If
    End If
    If m_prm.bUseXlsx Then
        m_wbXlsx = New XSSFWorkbook ' Excel 2007
        m_shXlsx = DirectCast(m_wbXlsx.CreateSheet(sTxtSheet), XSSFSheet)
        m_shXlsx.CreateFreezePane(prm.iNbFrozenColumns, 1)
        If m_prm.bCreateStandardSheet Then
            m_shStdrXlsx = DirectCast(m_wbXlsx.CreateSheet(sStdrSheet), XSSFSheet)
            m_shStdrXlsx.CreateFreezePane(prm.iNbFrozenColumns, 1)
            SetNumericStyle(m_numericCellStyleXlsx, bExcel2007:=True)
        End If
    End If

    delegMsg.ShowMsg("Filling workbook...")

    m_bAlertLineMax = False : m_bAlertColumnMax = False : m_bAlertCellTextLengthMax = False

    ' From 10 Mb read line by line
    Dim lFileSize& = New IO.FileInfo(sPath).Length
    Dim bLineByLine As Boolean = False
    If lFileSize > iBigFileSizeMb Then bLineByLine = True
    If Not bReadFileGeneric(sFieldDelimiter, bHeader, sPath, m_lineDeleg, delegMsg, _
        iNbLines, iNbColumns, encod:=encod, bLineByLine:=bLineByLine) Then Return False

    Dim bRes = bWrite()
    Return bRes

End Function

Public Sub FreeMemory()

    m_wb = Nothing : m_sh = Nothing
    m_wbXlsx = Nothing : m_shXlsx = Nothing

End Sub

Private Sub SetNumericStyle(ByRef style As ICellStyle, bExcel2007 As Boolean)

    Dim format As IDataFormat
    If bExcel2007 Then
        style = m_wbXlsx.CreateCellStyle
        format = m_wbXlsx.CreateDataFormat()
    Else
        style = m_wb.CreateCellStyle
        format = m_wb.CreateDataFormat()
    End If
    style.DataFormat = format.GetFormat("0")

End Sub

Public Function bWrite() As Boolean

    Dim sPath$ = Nothing

    m_delegMsg.ShowLongMsg("")
    m_delegMsg.ShowMsg("Checking Excel file...")
    Dim iNumColMax% = m_iNbFilledColMaxFound - 1
    If m_bXlsx Then
        sPath = m_sDestPathXlsx
    ElseIf m_prm.bUseXls Then
        sPath = m_sDestPathXls
    End If
    If Not bFileIsWritable(sPath, bNonExistentOk:=True, bPromptRetry:=True) Then Return False

    m_delegMsg.ShowMsg("Checking columns type...")
    If m_bXlsx Then
        If iNumColMax > iNbColMaxExcel2007 - 1 Then iNumColMax = iNbColMaxExcel2007 - 1
        SetWorkbookStyle(iNumColMax, bExcel2007:=True)
    ElseIf m_prm.bUseXls Then
        If iNumColMax > iNbColMaxAutoFilterExcel2003NPOI - 1 Then _
            iNumColMax = iNbColMaxAutoFilterExcel2003NPOI - 1
        SetWorkbookStyle(iNumColMax, bExcel2007:=False)
    End If

    m_delegMsg.ShowMsg("Writing workbook " & IO.Path.GetFileName(sPath) & "...")
    Try
        Using fs = New IO.FileStream(sPath, IO.FileMode.Create, IO.FileAccess.Write)
            If m_bXlsx Then
                m_wbXlsx.Write(fs)
            Else
                m_wb.Write(fs)
            End If
        End Using
    Catch ex As Exception
        m_delegMsg.ShowMsg("Error : Can't write the workbok !")
        ShowErrorMsg(ex, "File2XL : writing workbook", "Can't write the file : " & _
            IO.Path.GetFileName(sPath) & vbCrLf & sPath, sPossibleErrCause)
        Return False
    End Try

    m_delegMsg.ShowMsg(sMsgDone)
    Return True

End Function

Private Sub SetWorkbookStyle(iNumColMax%, bExcel2007 As Boolean)

    If iNumColMax < 0 Then Exit Sub

    ' Set header to gray
    Dim range As New CellRangeAddress(0, 0, 0, iNumColMax)
    Dim row0 As IRow
    Dim iNbColMax%
    If bExcel2007 Then
        m_shXlsx.SetAutoFilter(range)
        row0 = m_shXlsx.GetRow(0)
        iNbColMax = iNbColMaxExcel2007
    Else
        m_sh.SetAutoFilter(range)
        row0 = m_sh.GetRow(0)
        iNbColMax = iNbColMaxExcel2003
    End If

    For iNumField1 As Integer = 0 To m_lstFields.Count - 1
        If iNumField1 > row0.Cells.Count - 1 AndAlso iNumField1 < iNbColMax - 1 Then _
            row0.CreateCell(iNumField1)
    Next

    SetRowColor(row0, HSSFColor.Grey25Percent.Index, bExcel2007)

    Dim iMinColumnWidth% = m_prm.iMinColumnWidth
    Dim iMaxColumnWidth% = m_prm.iMaxColumnWidth
    Const iDisplayRate = 10
    Dim iNumField0% = 0
    Dim iNbFields0% = row0.Cells.Count
    For Each field In m_lstFields
        iNumField0 += 1
        If field.sType.StartsWith(clsFieldType.sNumeric, StringComparison.Ordinal) Then
            ' Color header
            If iNumField0 <= iNbFields0 Then

                SetCellColor(row0.Cells(iNumField0 - 1), HSSFColor.BrightGreen.Index, bExcel2007)

                If m_prm.bAutosizeColumns Then

                    If iNumField0 Mod iDisplayRate = 0 OrElse iNumField0 = iNbFields0 Then
                        m_delegMsg.ShowMsg("Text sheet : Autosizing column n°" & iNumField0 & "...")
                        If m_delegMsg.m_bCancel Then m_delegMsg.m_bCancel = False : Exit For
                    End If

                    If m_delegMsg.m_bCancel Then m_delegMsg.m_bCancel = False : Exit For

                    ' Set same column width on text sheet
                    If bExcel2007 Then
                        m_shXlsx.AutoSizeColumn(iNumField0 - 1) ' AutoFit

                        ' 20/05/2017
                        Dim iColWTxt% = m_shXlsx.GetColumnWidth(iNumField0 - 1)
                        If iColWTxt < iMinColumnWidth Then
                            iColWTxt = iMinColumnWidth
                            m_shXlsx.SetColumnWidth(iNumField0 - 1, iColWTxt)
                        End If
                        If iColWTxt > iMaxColumnWidth Then
                            iColWTxt = iMaxColumnWidth
                            m_shXlsx.SetColumnWidth(iNumField0 - 1, iColWTxt)
                        End If

                        If m_prm.bCreateStandardSheet Then _
                            m_shStdrXlsx.SetColumnWidth(iNumField0 - 1, iColWTxt)
                    Else
                        m_sh.AutoSizeColumn(iNumField0 - 1) ' AutoFit

                        ' 20/05/2017
                        Dim iColWTxt% = m_sh.GetColumnWidth(iNumField0 - 1)
                        'Debug.WriteLine("iColWTxt(" & iNumField0 & ")=" & iColWTxt)
                        If iColWTxt < iMinColumnWidth Then
                            iColWTxt = iMinColumnWidth
                            m_sh.SetColumnWidth(iNumField0 - 1, iColWTxt)
                        End If
                        If iColWTxt > iMaxColumnWidth Then
                            iColWTxt = iMaxColumnWidth
                            m_sh.SetColumnWidth(iNumField0 - 1, iColWTxt)
                        End If

                        If m_prm.bCreateStandardSheet Then _
                            m_shStdr.SetColumnWidth(iNumField0 - 1, iColWTxt)
                    End If

                End If

            End If
        End If
    Next

    If m_prm.bCreateStandardSheet Then

        If bExcel2007 Then
            m_shStdrXlsx.SetAutoFilter(range)
            row0 = m_shStdrXlsx.GetRow(0)
        Else
            m_shStdr.SetAutoFilter(range)
            row0 = m_shStdr.GetRow(0)
        End If

        For iNumField1 As Integer = 0 To m_lstFields.Count - 1
            If iNumField1 > row0.Cells.Count - 1 AndAlso iNumField1 < iNbColMax - 1 Then _
                row0.CreateCell(iNumField1)
        Next
        SetRowColor(row0, HSSFColor.Grey25Percent.Index, bExcel2007)

        Dim iNumField% = 0
        Dim iNbFields% = row0.Cells.Count
        For Each field In m_lstFields
            Dim iMemNumField% = iNumField
            iNumField += 1
            If field.sType.StartsWith(clsFieldType.sNumeric, StringComparison.Ordinal) Then
                ' Color header
                If iNumField <= iNbFields Then
                    SetCellColor(row0.Cells(iMemNumField), HSSFColor.BrightGreen.Index, bExcel2007)
                    If m_prm.bAutosizeColumns Then

                        If iNumField Mod iDisplayRate = 0 OrElse iNumField = iNbFields Then
                            m_delegMsg.ShowMsg("Standard sheet : Autosizing column n°" & iNumField & "...")
                            If m_delegMsg.m_bCancel Then m_delegMsg.m_bCancel = False : Exit For
                        End If

                        ' Set same column width on text sheet
                        If bExcel2007 Then
                            m_shStdrXlsx.AutoSizeColumn(iMemNumField) ' AutoFit
                            Dim iColWStdr% = m_shStdrXlsx.GetColumnWidth(iMemNumField)
                            Dim iColWTxt% = m_shXlsx.GetColumnWidth(iMemNumField)
                            Dim bResizeStdr As Boolean = False
                            Dim bResizeTxt As Boolean = False
                            If iColWStdr > iMaxColumnWidth Then iColWStdr = iMaxColumnWidth : bResizeStdr = True
                            If iColWTxt > iMaxColumnWidth Then iColWTxt = iMaxColumnWidth : bResizeTxt = True
                            If iColWStdr < iMinColumnWidth Then iColWStdr = iMinColumnWidth : bResizeStdr = True
                            If iColWTxt < iMinColumnWidth Then iColWTxt = iMinColumnWidth : bResizeTxt = True
                            If iColWTxt < iColWStdr Then
                                'm_shXlsx.SetColumnWidth(iMemNumField, iColWStdr)
                                iColWTxt = iColWStdr
                                bResizeTxt = True
                            ElseIf iColWTxt > iColWStdr Then
                                'm_shStdrXlsx.SetColumnWidth(iMemNumField, iColWTxt)
                                iColWStdr = iColWTxt
                                bResizeStdr = True
                            End If
                            If bResizeTxt Then m_shXlsx.SetColumnWidth(iMemNumField, iColWTxt)
                            If bResizeStdr Then m_shStdrXlsx.SetColumnWidth(iMemNumField, iColWStdr)
                        Else
                            m_shStdr.AutoSizeColumn(iMemNumField) ' AutoFit
                            Dim iColWStdr% = m_shStdr.GetColumnWidth(iMemNumField)
                            Dim iColWTxt% = m_sh.GetColumnWidth(iMemNumField)
                            'Debug.WriteLine("iColWStdr(" & iNumField0 & ")=" & iColWStdr)
                            'Debug.WriteLine("iColWTxt(" & iNumField0 & ")=" & iColWTxt)
                            Dim bResizeStdr As Boolean = False
                            Dim bResizeTxt As Boolean = False
                            If iColWStdr > iMaxColumnWidth Then iColWStdr = iMaxColumnWidth : bResizeStdr = True
                            If iColWTxt > iMaxColumnWidth Then iColWTxt = iMaxColumnWidth : bResizeTxt = True
                            If iColWStdr < iMinColumnWidth Then iColWStdr = iMinColumnWidth : bResizeStdr = True
                            If iColWTxt < iMinColumnWidth Then iColWTxt = iMinColumnWidth : bResizeTxt = True
                            If iColWTxt < iColWStdr Then
                                'm_sh.SetColumnWidth(iMemNumField, iColWStdr)
                                iColWTxt = iColWStdr
                                bResizeTxt = True
                            ElseIf iColWTxt > iColWStdr Then
                                'm_shStdr.SetColumnWidth(iMemNumField, iColWTxt)
                                iColWStdr = iColWTxt
                                bResizeStdr = True
                            End If
                            If bResizeTxt Then m_sh.SetColumnWidth(iMemNumField, iColWTxt)
                            If bResizeStdr Then m_shStdr.SetColumnWidth(iMemNumField, iColWStdr)
                        End If

                    End If
                End If
            End If
        Next

        If bExcel2007 Then
            m_wbXlsx.SetSelectedTab(1)
            m_wbXlsx.SetActiveSheet(1)
        Else
            m_wb.SetSelectedTab(1)
            m_wb.SetActiveSheet(1)
        End If

    End If

End Sub

Private Sub UpdateMsg()

    Dim iNbColMaxExcel% = iNbColMaxExcel2003
    Dim iNbLinesMaxExcel% = iNbLinesMaxExcel2003
    If m_bXlsx Then
        iNbColMaxExcel = iNbColMaxExcel2007
        iNbLinesMaxExcel = iNbLinesMaxExcel2007
    End If
    m_sMsgMaxColumns = "The number of columns exceeds the maximum allowed (" & _
        iNbColMaxExcel & ")." & vbLf & "Next columns will be ignored."
    m_sMsgMaxLines = "The number of lines exceeds the maximum allowed (" & _
        iNbLinesMaxExcel & ")." & vbLf & "Next lines will be ignored."

End Sub

Private Sub NewSplitLine(sender As Object, e As clsSplitLineEventArgs) _
    Handles m_lineDeleg.EvNewSplitLine

    If m_bDetectColumnType Then
        Dim lstFields = New List(Of String)
        For Each sField In e.m_asFields
            lstFields.Add(sField)
        Next
        m_splitLines.Add(lstFields)
        Exit Sub
    End If

    ' Fill Excel workbook

    If m_iNumLine >= iNbLinesMaxExcel2003 Then
        If Not m_bAlertLineMax Then
            If Not m_prm.bUseXlsx AndAlso m_prm.bUseXls Then
                Dim row0 = m_sh.GetRow(iNbLinesMaxExcel2003 - 1)
                AlerteRow(row0, bExcel2007:=False)
                Dim row1 = m_shStdr.GetRow(iNbLinesMaxExcel2003 - 1)
                AlerteRow(row1, bExcel2007:=False)
            Else
                m_bXlsx = True : UpdateMsg()
            End If
        End If
        If Not m_prm.bUseXlsx Then Exit Sub
    End If

    If m_iNumLine >= iNbLinesMaxExcel2007 Then
        If Not m_bAlertLineMax Then
            Dim row0 = m_shXlsx.GetRow(iNbLinesMaxExcel2007 - 1)
            AlerteRow(row0, bExcel2007:=True)
            Dim row1 = m_shStdrXlsx.GetRow(iNbLinesMaxExcel2007 - 1)
            AlerteRow(row1, bExcel2007:=True)
        End If
        Exit Sub
    End If

    Dim iNbCol% = e.m_asFields.Count
    If iNbCol > m_iNbColMaxFound Then m_iNbColMaxFound = iNbCol

    Dim row As IRow
    If m_prm.bUseXlsx Then
        row = m_shXlsx.CreateRow(m_iNumLine)
        WriteRow(row, e.m_asFields, iNbColMaxExcel2007, bExcel2007:=True)
        If m_prm.bCreateStandardSheet Then
            row = m_shStdrXlsx.CreateRow(m_iNumLine)
            WriteRow(row, e.m_asFields, iNbColMaxExcel2007, bExcel2007:=True, bConv:=True)
        End If
    End If
    If m_prm.bUseXls AndAlso Not m_bXlsx Then
        row = m_sh.CreateRow(m_iNumLine)
        WriteRow(row, e.m_asFields, iNbColMaxExcel2003, bExcel2007:=False)
        If m_prm.bCreateStandardSheet Then
            row = m_shStdr.CreateRow(m_iNumLine)
            WriteRow(row, e.m_asFields, iNbColMaxExcel2003, bExcel2007:=False, bConv:=True)
        End If
    End If

    m_iNumLine += 1

End Sub

Private Sub AlerteRow(row0 As IRow, bExcel2007 As Boolean)

    'Dim val = row0.Cells(0)
    'Dim sCellVal$ = " " & val.StringCellValue ' 05/06/2016 Exception with NPOI 2.2.1.0 !
    'Dim sCellVal$ = " " & val.RichStringCellValue.String ' Idem
    Dim sCellVal$ = "" ' 05/06/2016
    SetCellValue(row0.Cells(0), sMsgNextLinesIgnored & sCellVal, bExcel2007)
    SetCellColor(row0.Cells(0), HSSFColor.Orange.Index, bExcel2007)

    If Not m_bAlertLineMax Then
        m_bAlertLineMax = True
        If m_prm.bMsgBoxAlert Then MsgBox(m_sMsgMaxLines, vbExclamation, m_sMsgTitle)
    End If

    m_delegMsg.m_bIgnoreNextLines = True

End Sub

Private Sub WriteRow(row As IRow, asFields$(), iNbColMaxExcel%, bExcel2007 As Boolean, _
    Optional bConv As Boolean = False)

    Dim iNumField% = 0

    For Each sField In asFields

        If m_prm.bUseXlsx AndAlso iNumField > iNbColMaxExcel2003 - 1 Then
            If Not m_bXlsx Then m_bXlsx = True : UpdateMsg()
            If Not bExcel2007 Then Exit Sub
        End If

        If iNumField > iNbColMaxExcel - 1 Then
            'Dim val = row.Cells(iNumField - 1)
            'Dim sCellVal$ = " " & val.StringCellValue ' 05/06/2016 Exception with NPOI 2.2.1.0 !
            'Dim sCellVal$ = " " & val.RichStringCellValue.String ' 05/06/2016 Idem
            Dim sCellVal$ = "" ' 05/06/2016
            SetCellValue(row.Cells(iNumField - 1), sMsgNextColumnsIgnored & sCellVal, bExcel2007)
            SetCellColor(row.Cells(iNumField - 1), HSSFColor.Orange.Index, bExcel2007)
            If m_prm.bMsgBoxAlert AndAlso Not m_bAlertColumnMax Then
                MsgBox(m_sMsgMaxColumns, vbExclamation, m_sMsgTitle)
                m_bAlertColumnMax = True
            End If
            Exit For
        End If

        row.CreateCell(iNumField)

        ' Remove NULL value only for the Standard sheet (bConv = True), not for the Text sheet (bConv = False)
        If m_prm.bRemoveNULL AndAlso bConv AndAlso sField = sNULL Then sField = "" ' 28/04/2017

        Dim bValue As Boolean = False
        If sField.Length > 0 Then bValue = True

        If bValue Then

            If bConv AndAlso iNumField < m_lstFields.Count Then
                Dim field = m_lstFields(iNumField)
                If field.sType = clsFieldType.sText Then
                    If row.RowNum = 0 Then
                        SetCellValue(row.Cells(iNumField), sField.Replace(sQuotes, sEmpty), bExcel2007)
                    Else
                        SetCellValue(row.Cells(iNumField), sField, bExcel2007)
                    End If
                ElseIf field.sType = clsFieldType.sTextWithQuotes Then
                    SetCellValue(row.Cells(iNumField), sField.Replace(sQuotes, sEmpty), bExcel2007)
                Else
                    Dim sFieldConv$
                    Select Case field.sType
                    Case clsFieldType.sNumeric : sFieldConv = sField
                    Case clsFieldType.sNumericC2P : sFieldConv = sField.Replace(sComma, sPeriod)
                    Case clsFieldType.sNumericP2C : sFieldConv = sField.Replace(sPeriod, sComma)
                    Case clsFieldType.sNumericWithQuotes : sFieldConv = sField.Replace(sQuotes, sEmpty)
                    Case clsFieldType.sNumericC2PWithQuotes : sFieldConv = sField.Replace(sComma, sPeriod).Replace(sQuotes, sEmpty)
                    Case clsFieldType.sNumericP2CWithQuotes : sFieldConv = sField.Replace(sPeriod, sComma).Replace(sQuotes, sEmpty)
                    Case Else : sFieldConv = sField
                    End Select

                    Dim iMult% = 1
                    If field.bCanEndWithMinus Then
                        Dim sFieldTrim$ = sField.Trim
                        If sFieldTrim.EndsWith("-", StringComparison.Ordinal) Then
                            sFieldConv = sFieldTrim.Substring(0, sFieldTrim.Length - 1)
                            iMult = -1
                        End If
                    End If

                    Dim bOk As Boolean
                    Dim rVal# = iMult * rFastConv(sFieldConv, , bOk)
                    If bOk Then
                        row.Cells(iNumField).SetCellValue(rVal)
                        If bExcel2007 Then
                            row.Cells(iNumField).CellStyle = m_numericCellStyleXlsx
                        Else
                            row.Cells(iNumField).CellStyle = m_numericCellStyleXls
                        End If
                    Else
                        ' Header fields
                        If field.sType.EndsWith(sPostFixWithQuotes, StringComparison.Ordinal) Then
                            SetCellValue(row.Cells(iNumField), sField.Replace(sQuotes, sEmpty), bExcel2007)
                        Else
                            SetCellValue(row.Cells(iNumField), sField, bExcel2007)
                        End If
                    End If
                End If
            Else
                SetCellValue(row.Cells(iNumField), sField, bExcel2007)
            End If

        End If
        iNumField += 1
        If bValue AndAlso iNumField > m_iNbFilledColMaxFound Then m_iNbFilledColMaxFound = iNumField
    Next

End Sub

Private Sub NewLine(sender As Object, e As clsLineEventArgs) Handles m_lineDeleg.EvNewLine
    m_lines.Add(e.sLine)
End Sub

' Attribute for inline to avoid function overhead
<MethodImpl(MethodImplOptions.AggressiveInlining)> _
Private Sub SetCellValue(cell As ICell, sValue$, bExcel2007 As Boolean)

    Const bReplaceTab As Boolean = False ' This constant may be a setting in a next release
    If bReplaceTab AndAlso sValue.IndexOf(vbTab) > -1 Then
        sValue = sValue.Replace(vbTab, "    ")
    End If

    If sValue.Length <= iNbCarMaxCell Then
        cell.SetCellValue(sValue)
    Else
        Dim iNbCar% = iNbCarMaxCell - sMsgNextCarIgnored.Length
        Dim sTruncVal$ = sMsgNextCarIgnored & sValue.Substring(0, iNbCar)
        cell.SetCellValue(sTruncVal)
        SetCellColor(cell, HSSFColor.Orange.Index, bExcel2007)
        If m_prm.bMsgBoxAlert AndAlso Not m_bAlertCellTextLengthMax Then
            MsgBox(sMsgMaxCarCell, vbExclamation, m_sMsgTitle)
            m_bAlertCellTextLengthMax = True
        End If
    End If

End Sub

Private Sub SetRowColor(row As IRow, indexColor As Short, bExcel2007 As Boolean)

    If IsNothing(row) Then
        If bDebug Then Stop
        Exit Sub
    End If

    Const iColMin% = 0
    Dim iColMax% = row.LastCellNum

    Dim style As ICellStyle
    If bExcel2007 Then
        style = m_wbXlsx.CreateCellStyle
    Else
        style = m_wb.CreateCellStyle
    End If

    style.FillForegroundColor = indexColor
    style.FillPattern = FillPattern.SolidForeground

    For i = iColMin To iColMax - 1
        Dim cell = row.GetCell(i)
        cell.CellStyle = style
    Next

End Sub

Private Sub SetCellColor(cell As ICell, indexColor As Short, bExcel2007 As Boolean)

    If IsNothing(cell) Then
        If bDebug Then Stop
        Exit Sub
    End If

    Dim style As ICellStyle
    If bExcel2007 Then
        style = m_wbXlsx.CreateCellStyle
    Else
        style = m_wb.CreateCellStyle
    End If

    style.FillForegroundColor = indexColor
    style.FillPattern = FillPattern.SolidForeground
    cell.CellStyle = style

End Sub

Private Sub FindProbDelimiter(sDelimiterList$, sDefaultDelimiter$, ByRef sFieldDelimiter$)

    Const bDebugSort As Boolean = False

    sFieldDelimiter = String.Empty

    Dim acDelimiters = sDelimiterList.ToCharArray

    Dim dicStat As New SortDic(Of String, clsOcc)
    For Each c In acDelimiters
        Dim s$ = c
        dicStat.Add(s, New clsOcc(s, 0, m_prm.bPreferMultipleDelimiter)) ' Count succes
    Next
    ' Count also "," and ";" if required
    ' 16/04/2017 AndAlso m_prm.bPreferMultipleDelimiter
    If m_prm.bUseQuotesCommaQuotesDelimiter AndAlso m_prm.bPreferMultipleDelimiter Then
        dicStat.Add(sQuotesCommaQuotesDelimiter, _
            New clsOcc(sQuotesCommaQuotesDelimiter, 0, m_prm.bPreferMultipleDelimiter))
        dicStat.Add(sQuotesSemiColonQuotesDelimiter, _
            New clsOcc(sQuotesSemiColonQuotesDelimiter, 0, m_prm.bPreferMultipleDelimiter))
    End If

    Const sSorting = "iWeight DESC, iNbOcc DESC, iOccLength DESC" ' Fields of clsOcc
    Dim iNumLine% = 0
    For Each sLine In m_lines
        iNumLine += 1
        Dim dic As New SortDic(Of String, clsOcc)
        For Each c In acDelimiters
            Dim s$ = c
            Dim iNbOcc% = iNbOccurrences(sLine, s)
            If dic.ContainsKey(s) Then Continue For
            dic.Add(s, New clsOcc(s, iNbOcc, m_prm.bPreferMultipleDelimiter))
        Next
        ' 16/04/2017 AndAlso m_prm.bPreferMultipleDelimiter
        If m_prm.bUseQuotesCommaQuotesDelimiter AndAlso m_prm.bPreferMultipleDelimiter Then
            Dim iNbOcc% = iNbOccurrences(sLine, sQuotesCommaQuotesDelimiter)
            If Not dic.ContainsKey(sQuotesCommaQuotesDelimiter) Then
                dic.Add(sQuotesCommaQuotesDelimiter, _
                    New clsOcc(sQuotesCommaQuotesDelimiter, iNbOcc, m_prm.bPreferMultipleDelimiter))
            End If
            Dim iNbOcc2% = iNbOccurrences(sLine, sQuotesSemiColonQuotesDelimiter)
            If Not dic.ContainsKey(sQuotesSemiColonQuotesDelimiter) Then
                dic.Add(sQuotesSemiColonQuotesDelimiter, _
                    New clsOcc(sQuotesSemiColonQuotesDelimiter, iNbOcc2, m_prm.bPreferMultipleDelimiter))
            End If
        End If
        If bDebugSort Then Debug.WriteLine("Result line n°" & iNumLine & " :")
        Dim iNumSep% = 0
        ' First sort by number of occurrences, then by occurrence length, so that "," can win against ,
        For Each occ In dic.Sort(sSorting)
            If bDebugSort Then Debug.WriteLine(occ.s & "=" & occ.iNbOcc & " (" & occ.iOccLength & " car.)")
            If iNumSep = 0 AndAlso occ.iNbOcc > 0 Then dicStat(occ.s).iNbOcc += 1
            iNumSep += 1
        Next
    Next

    If bDebugSort Then
        Debug.WriteLine("")
        Debug.WriteLine("")
        Debug.WriteLine("Results :")
    End If

    Dim sProb$ = String.Empty
    Dim iNumSep2% = 0
    For Each occ In dicStat.Sort(sSorting)
        If bDebugSort Then Debug.WriteLine(occ.s & "=" & occ.iNbOcc & " wins / " & m_lines.Count)
        If iNumSep2 = 0 AndAlso occ.iNbOcc > 0 Then sProb = occ.s ' Keep the winner
        iNumSep2 += 1
    Next

    If sProb = sQuotesCommaQuotesDelimiter OrElse sProb = sQuotesSemiColonQuotesDelimiter Then
        sFieldDelimiter = sProb
    Else
        If sProb = String.Empty Then
            If bDebugSort Then Debug.WriteLine("No delimiter found")
            If m_prm.bAlertForNoDelimiterFound Then
                Dim sMsg$ = "No delimiter found !"
                If Not String.IsNullOrEmpty(sDefaultDelimiter) Then
                    sMsg &= vbLf & "Default delimiter will be use : [" & sDefaultDelimiter & "]"
                End If
                MsgBox(sMsg, MsgBoxStyle.Exclamation, m_sMsgTitle)
            End If
            sProb = sDefaultDelimiter
        End If
        sFieldDelimiter = sProb
    End If

End Sub

Private Sub FindColumnsType(ByRef lstFields As List(Of clsField), ByRef bOnlyTextFields As Boolean)

    Const bDebugColType As Boolean = False

    bOnlyTextFields = True

    Dim lstFields0 As New List(Of SortDic(Of String, clsField))
    Dim lstNameOfFields As New List(Of String)
    Dim lstMinusExistsForFields As New List(Of Boolean)

    Dim iNumLine% = 0
    For Each sLine In m_splitLines
        If bDebugColType Then Debug.WriteLine("iNumLine=" & iNumLine + 1)
        Dim iNumField% = 0
        For Each sField In sLine

            ' 28/04/2017
            If m_prm.bRemoveNULL AndAlso sField = sNULL Then
                Continue For ' Do not count null value
            End If

            'Dim sFieldMinus = ""
            Dim sFieldTrim$ = sField.Trim
            Dim bEndWithMinus As Boolean = False
            If sFieldTrim.EndsWith("-", StringComparison.Ordinal) Then
                bEndWithMinus = True
                'sFieldMinus = sFieldTrim.Substring(0, sFieldTrim.Length - 1)
            End If

            Dim dic As SortDic(Of String, clsField)
            Dim sFieldName$ = ""
            If iNumLine = 0 OrElse iNumField >= lstFields0.Count Then
                dic = New SortDic(Of String, clsField)
                lstFields0.Add(dic)

                If iNumLine = 0 Then sFieldName = sField
                If sFieldName.Trim.Length = 0 Then sFieldName = "Field n°" & iNumField + 1
                lstNameOfFields.Add(sFieldName)
                lstMinusExistsForFields.Add(bEndWithMinus)

            Else
                dic = lstFields0(iNumField)
                sFieldName = lstNameOfFields(iNumField)

                Dim bEndWithMinus0 = lstMinusExistsForFields(iNumField)
                If bEndWithMinus AndAlso Not bEndWithMinus0 Then
                    lstMinusExistsForFields(iNumField) = bEndWithMinus
                End If
            End If

            If IsNumeric(sField) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumeric)
            ElseIf IsNumeric(sField.Replace(sPeriod, sComma)) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumericP2C)
            ElseIf IsNumeric(sField.Replace(sComma, sPeriod)) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumericC2P)
            ElseIf IsNumeric(sField.Replace(sQuotes, sEmpty)) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumericWithQuotes)
            ElseIf IsNumeric(sField.Replace(sPeriod, sComma).Replace(sQuotes, sEmpty)) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumericP2CWithQuotes)
            ElseIf IsNumeric(sField.Replace(sComma, sPeriod).Replace(sQuotes, sEmpty)) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sNumericC2PWithQuotes)
            ElseIf sField.Contains(sQuotes) Then
                AddField(dic, iNumField, sFieldName, clsFieldType.sTextWithQuotes)
            Else
                AddField(dic, iNumField, sFieldName, clsFieldType.sText)
            End If

            If bDebugColType Then Debug.WriteLine(sFieldName & "=" & sField)
            iNumField += 1
        Next
        iNumLine += 1
    Next

    lstFields = New List(Of clsField)
    Dim iNumField2% = 0
    For Each dic In lstFields0
        Dim iNumSep2% = 0
        For Each field In dic.Sort("iNbOcc DESC")
            If bDebugColType Then Debug.WriteLine(field.iNumField & " : " & field.sField & _
                ", " & field.sType & ", iNbOcc=" & field.iNbOcc)
            ' Keep the max.
            If iNumSep2 = 0 Then
                lstFields.Add(field)
                If field.sType <> clsFieldType.sText Then bOnlyTextFields = False
            End If
            iNumSep2 += 1
        Next
        iNumField2 += 1
    Next

    Dim iNumField3% = 0
    For Each field In lstFields
        field.bCanEndWithMinus = lstMinusExistsForFields(iNumField3)
        iNumField3 += 1
    Next

End Sub

Private Shared Sub AddField(dic As SortDic(Of String, clsField), iNumField%, sFieldName$, sFieldType$)

    If Not dic.ContainsKey(sFieldType) Then
        Dim field = New clsField(iNumField, sFieldName, sFieldType)
        dic.Add(sFieldType, field)
    Else
        Dim field = dic(sFieldType)
        field.iNbOcc += 1
    End If

End Sub

End Class