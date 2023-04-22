
' File modGenRead.vb : Generic text reading module
' ------------------

Imports System.Text ' StringBuilder
Imports System.IO

Module modGenericReading

    ' Read big files (> 10 Mb) line by line
    Public Const iBigFileSizeMb% = 10 * 1024 * 1024 ' 10 Mb

#Region "Delegate (call-back)"

    Public Class clsLineEventArgs : Inherits EventArgs
        Private m_sLine$ = ""
        Public Sub New(sLine$)
            If String.IsNullOrEmpty(sLine) Then sLine = ""
            Me.m_sLine = sLine
        End Sub
        Public ReadOnly Property sLine$()
            Get
                Return Me.m_sLine
            End Get
        End Property
    End Class

    Public Class clsSplitLineEventArgs : Inherits EventArgs

        Public m_asFields$()
        Public m_iNbColumns% = 0

        Public Sub New(sLine$, sFieldDelimiter$, bQuotesDelimiter As Boolean)

            ' This constructor works only using "," or ";"

            If String.IsNullOrEmpty(sLine) Then sLine = ""

            ' Split using " : remove first and last ones
            If bQuotesDelimiter AndAlso Not String.IsNullOrEmpty(sLine) AndAlso sLine.Length > 2 Then
                sLine = sLine.Substring(1, sLine.Length - 2)
            End If

            If String.IsNullOrEmpty(sFieldDelimiter) Then
                ReDim Me.m_asFields(0)
                Me.m_asFields(0) = sLine
                m_iNbColumns = 1
            Else
                Me.m_asFields = Split(sLine, sFieldDelimiter)
                m_iNbColumns = Me.m_asFields.GetUpperBound(0)
            End If

        End Sub

    End Class

    Public Class clsDelegLine

        Public Delegate Sub EvHandlerLine(sender As Object, e As clsLineEventArgs)
        Public Event EvNewLine As EvHandlerLine

        Public Delegate Sub EvHandlerSplitLine(sender As Object, e As clsSplitLineEventArgs)
        Public Event EvNewSplitLine As EvHandlerSplitLine

        Public Sub NewLine(sLine$)
            Dim e As New clsLineEventArgs(sLine)
            RaiseEvent EvNewLine(Me, e)
        End Sub

        Public Sub NewSplitLine(sLine$, sFieldDelimiter$, bQuotesDelimiter As Boolean, ByRef iNbColumns%)
            Dim e As New clsSplitLineEventArgs(sLine, sFieldDelimiter, bQuotesDelimiter)
            RaiseEvent EvNewSplitLine(Me, e)
            iNbColumns = e.m_iNbColumns
        End Sub

    End Class

#End Region

    Public Function bReadFileGenericSmart(sFieldDelimiter$, bHeader As Boolean,
        sPath$, delegLine As clsDelegLine, msgDeleg As clsDelegMsg,
        ByRef iNbLines%, ByRef iNbColumns%) As Boolean

        Dim encod = GetEncoding(sPath)
        ' If encoding is ASCII, set the Latin alphabet to preserve for example accents
        ' Default = System.Text.SBCSCodePageEncoding = Encoding.GetEncoding(1252)
        If encod Is Encoding.ASCII Then encod = Encoding.Default

        ' From 10 Mb read line by line
        Dim lTailleFic& = New IO.FileInfo(sPath).Length
        Dim bLineByLineMode As Boolean = False
        If lTailleFic > iBigFileSizeMb Then bLineByLineMode = True

        Return bReadFileGeneric(sFieldDelimiter, bHeader, sPath, delegLine, msgDeleg,
            iNbLines, iNbColumns, bLineByLineMode, encod:=encod)

    End Function

    Public Function bReadFileGeneric(sFieldDelimiter$, bHeader As Boolean,
        sPath$, lineDeleg As clsDelegLine, msgDeleg As clsDelegMsg,
        ByRef iNbLines%, ByRef iNbColumns%,
        Optional bLineByLine As Boolean = False,
        Optional bOnlyFirstLines As Boolean = False,
        Optional bOnlyFirstSplitLines As Boolean = False,
        Optional encod As Encoding = Nothing,
        Optional iNbLinesAnalyzed% = 10) As Boolean

        iNbLines = -1 ' -1 = Not started

        Dim sFile$ = IO.Path.GetFileName(sPath)
        Dim sMsg0$ = "Loading " & sFile & "..."
        msgDeleg.ShowMsg(sMsg0)
        Dim sMsg1$ = "Loading..." & vbLf & sPath
        msgDeleg.ShowLongMsg(sMsg1)

        If Not bFileExists(sPath, bPrompt:=True) Then Return False

        Dim bQuotesDelimiter As Boolean = False
        If sFieldDelimiter = sQuotesCommaQuotesDelimiter OrElse
           sFieldDelimiter = sQuotesSemiColonQuotesDelimiter Then bQuotesDelimiter = True

        If IsNothing(encod) Then encod = Encoding.Default

        Dim iNumLine% = 0
        Dim iDisplayRate0% = iDisplayRate
        If iNbColumns > 0 Then
            Select Case iNbColumns
                Case 1 To 5 : iDisplayRate0 = 10000
                Case 6 To 10 : iDisplayRate0 = iDisplayRate
                Case 11 To 50 : iDisplayRate0 = 500
                Case 51 To 100 : iDisplayRate0 = 100
                Case 101 To 1000 : iDisplayRate0 = 10
                Case Else
                    iDisplayRate0 = 1
            End Select
        End If

        If bLineByLine OrElse bOnlyFirstLines OrElse bOnlyFirstSplitLines Then

            ' Read line by line
            Dim fs As FileStream = Nothing
            Try
                Dim ci = Globalization.CultureInfo.CurrentCulture()
                Dim lFileSize& = New IO.FileInfo(sPath).Length
                Dim share As IO.FileShare = IO.FileShare.ReadWrite
                fs = New IO.FileStream(sPath, IO.FileMode.Open, IO.FileAccess.Read, share)
                Dim lPosition& = 0
                Using sr As New IO.StreamReader(fs, encod)
                    fs = Nothing ' 19/05/2017 Do not use fs.Position inside this loop
                    Do
                        Dim sLine$ = sr.ReadLine()
                        ' 20/08/2017 If Not String.IsNullOrEmpty(sLine) Then 
                        If Not String.IsNullOrEmpty(sLine) Then lPosition += sLine.Length
                        iNumLine += 1

                        If bOnlyFirstLines Then
                            If iNumLine > iNbLinesAnalyzed Then Return True
                            If IsNothing(sLine) Then Continue Do
                            lineDeleg.NewLine(sLine)
                            Continue Do
                        End If

                        If bHeader AndAlso Not bOnlyFirstLines AndAlso iNumLine = 1 Then Continue Do ' Header

                        If msgDeleg.m_bIgnoreNextLines Then Exit Do
                        If IsNothing(sLine) Then Continue Do

                        Dim iNbColumns0% = 0
                        lineDeleg.NewSplitLine(sLine, sFieldDelimiter, bQuotesDelimiter, iNbColumns0)
                        If iNbColumns0 > iNbColumns Then iNbColumns = iNbColumns0

                        If bOnlyFirstSplitLines Then
                            If iNumLine > iNbLinesAnalyzed Then Return True
                            Continue Do
                        End If

                        If iNumLine Mod iDisplayRate0 = 0 Then
                            'Dim lFilePos& = fs.Position
                            Dim lFilePos& = lPosition ' 19/05/2017
                            Dim rPC! = 100 * CSng(lFilePos / lFileSize)
                            Dim sPC$ = iNumLine & " (" & rPC.ToString("0.00", ci) & " %)..."
                            Dim sMsg$ = sFile & " lines : " & sPC
                            Dim sLongMsg$ = sPC & vbLf & sPath & vbLf & sRAMInfo()
                            msgDeleg.ShowMsg("Loading : " & sMsg)
                            msgDeleg.ShowLongMsg("Loading : " & sLongMsg)
                            WaitPause(msgDeleg, "Paused : " & sMsg, "Paused : " & sLongMsg)
                            If msgDeleg.m_bCancel Then Return False
                        End If
                    Loop While Not sr.EndOfStream
                End Using
                If Not msgDeleg.m_bIgnoreNextLines Then
                    iNbLines = iNumLine
                    Dim sPC1$ = iNumLine & " (" & (100).ToString("0.00", ci) & " %)"
                    Dim sMsg$ = "Loading " & sFile & " lines : " & sPC1
                    Dim sLongMsg$ = "Loading : " & sPC1 & vbLf & sPath & vbLf & sRAMInfo()
                    msgDeleg.ShowMsg(sMsg)
                    msgDeleg.ShowLongMsg(sLongMsg)
                End If
            Catch ex As Exception
                Throw
                Return False
            Finally
                ' 19/05/2017 Right code to suppress CA2202 warning, but fs.Position
                '  cannot be read inside the loop
                If fs IsNot Nothing Then fs.Dispose() ' CA2000
            End Try

        Else

            ' Read whole file
            Dim asLines$() = asReadFile(sPath, bReadOnly:=True, encod:=encod)
            If IsNothing(asLines) Then Return False
            iNbLines = asLines.Count
            For Each sLine As String In asLines
                iNumLine += 1
                If bHeader AndAlso iNumLine = 1 Then Continue For ' Header
                If msgDeleg.m_bIgnoreNextLines Then Exit For

                Dim iNbColumns0% = 0
                lineDeleg.NewSplitLine(sLine, sFieldDelimiter, bQuotesDelimiter, iNbColumns0)
                If iNbColumns0 > iNbColumns Then iNbColumns = iNbColumns0

                If iNumLine Mod iDisplayRate0 = 0 OrElse iNumLine = iNbLines Then
                    Dim sProgress$ = iNumLine & "/" & iNbLines
                    Dim sMsg$ = sFile & " lines... " & sProgress
                    Dim sLongMsg$ = sProgress & vbLf & sPath & vbLf & sRAMInfo()
                    msgDeleg.ShowMsg("Loading : " & sMsg)
                    msgDeleg.ShowLongMsg("Loading : " & sLongMsg)
                    WaitPause(msgDeleg, "Paused : " & sMsg, "Paused : " & sLongMsg)
                    If msgDeleg.m_bCancel Then Return False
                End If
            Next

        End If

        If bHeader AndAlso iNbLines > 0 Then iNbLines -= 1

        msgDeleg.ShowMsg(sMsgDone)
        msgDeleg.ShowLongMsg(sMsgDone)

        Return True

    End Function

    Private Sub WaitPause(msgDeleg As clsDelegMsg, sMsgPause$, sLongMsgPause$)
        While msgDeleg.m_bPause
            msgDeleg.ShowMsg(sMsgPause)
            msgDeleg.ShowLongMsg(sLongMsgPause)
            If msgDeleg.m_bCancel Then Exit Sub
            Threading.Thread.Sleep(500)
        End While
    End Sub

End Module