
' File2XL : Open a csv file into MS-Excel with pre-formatted cells
' ----------------------------------------------------------------
' Documentation : File2XL.html
' http://patrice.dargenton.free.fr/CodesSources/File2XL.html
' http://patrice.dargenton.free.fr/CodesSources/File2XL.vbproj.html
' Version 1.07 - 22/04/2023
' By Patrice Dargenton : mailto:patrice.dargenton@free.fr
' http://patrice.dargenton.free.fr/index.html
' http://patrice.dargenton.free.fr/CodesSources/index.html
' ----------------------------------------------------------------

' Naming convention :
' -----------------
' b for Boolean (True or False)
' i for Integer : %
' l for Long : &
' r for Real number (Single!, Double# or Decimal : D)
' s for String : $
' c for Char or Byte
' d for Date
' u for Unsigned (positif integer)
' a for Array : ()
' o for Object
' m_ for member variable of a class or of a form (but not for constants)
' frm for Form
' cls for Class
' mod for Module
' ...
' -----------------

' File frmFile2XL.vb : Main form
' ------------------

Imports System.Text ' for StringBuilder

Public Class frmFile2XL

    Private Const bDelWorkBookOnCloseDef As Boolean = True
    Private m_bDelWorkBookOnClose As Boolean = bDelWorkBookOnCloseDef

    Private Const sContextMenu_FileTypeAll$ = "*" ' Every file (every text or csv file to open in Excel)
    Private Const sContextMenu_CmdKeyOpen$ = "File2XL.Open"
    Private Const sContextMenu_CmdKeyOpenDescr$ = "Open in MS-Excel using File2XL"
    Private Const sContextMenu_CmdKeyOpen2$ = "File2XL.Open2"
    Private Const sContextMenu_CmdKeyOpen2Descr$ = "Open in MS-Excel using File2XL (single delimiter)"
    Private Const sSingleDelimiterArg$ = "SingleDelimiter" ' For example , rather than ","

    Private WithEvents m_delegMsg As New clsDelegMsg
    Private m_bInit As Boolean = False
    Private m_bXlsExists, m_bXlsxExists As Boolean
    Private m_bClosing As Boolean = False

    Private m_f2xl As New clsFile2XL

    Private Sub frmFile2XL_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Not m_bInit Then
            m_bInit = True
            Initialization()
        End If
    End Sub

    Private Sub frmFile2XL_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        m_delegMsg.m_bCancel = True
        Quit()
    End Sub

    Private Sub Initialization()

        SetMsgTitle(sMsgTitle)
        Dim sTxt$ = sMsgTitle & " " & sAppVersion & " (" & sAppDate & ")"
        If bDebug Then sTxt &= " - Debug"
        'If Not is64BitProcess() Then sTxt &= " - 32 bits"
        Me.Text = sTxt

        Me.cmdCancel.Visible = False
        Me.cmdStart.Visible = False
        Me.cmdShow.Visible = False
        If bRelease Then Me.cmdCreateTestFiles.Visible = False

        Me.ToolTip1.SetToolTip(Me.cmdAddContextMenu,
            "Add context menu to open files in Excel using File2XL " &
            "(this requires administrator privileges, run as admin. File2XL for this operation)")
        Me.ToolTip1.SetToolTip(Me.cmdRemoveContextMenu,
            "Remove context menu for opening files in Excel using File2XL " &
            "(this requires administrator privileges, run as admin. File2XL for this operation)")

        CheckContextMenu()

        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command
        'MsgBox("File2XL : " & sArg0)
        If bDebug Then
            sArg0 = Application.StartupPath & "\Tmp\Test256Col.dat"
            'sArg0 = Application.StartupPath & "\Tmp\Standard_sheet.csv"
            sArg0 = Application.StartupPath & "\Tmp\Test.csv"
        Else
            Me.cmdCreateTestFiles.Visible = False
        End If

        If sArg0.Length > 0 Then
            Dim asArgs$() = asCmdLineArg(sArg0)
            If asArgs.Length > 0 Then
                Dim sArgument$ = asArgs(0)
                Dim sArgument2$ = ""
                Dim bSingleDelimiter As Boolean = False
                If asArgs.Length > 1 Then
                    sArgument2 = asArgs(1)
                    If sArgument2 = sSingleDelimiterArg Then bSingleDelimiter = True
                End If
                ShowButtons()
                HideContextMenus()
                Activation(bActivate:=False)
                ShowMessage("Starting...")
                If bStart(sArgument, bSingleDelimiter, sArgument2) Then
                    Activation(bActivate:=True)
                    Quit()
                Else
                    ShowMessage("Error !")
                    Activation(bActivate:=True)
                    Me.cmdShow.Enabled = False
                End If
                Exit Sub
            End If
        End If

EndSub:
        ShowMessage("Ready.")

    End Sub

    Private Function bStart(sPath$, bSingleDelimiter As Boolean, sArgument$) As Boolean

        If Not bFileExists(sPath, bPrompt:=True) Then Return False

        Dim sFileName$ = IO.Path.GetDirectoryName(sPath) & "\" &
            IO.Path.GetFileNameWithoutExtension(sPath)
        Dim sPostFix$ = ""
        If bSingleDelimiter Then sPostFix = "_" & sSingleDelimiterArg
        m_f2xl.m_sDestPathXls = sFileName & sPostFix & ".xls"
        m_f2xl.m_sDestPathXlsx = sFileName & sPostFix & ".xlsx"
        m_bXlsExists = bFileExists(m_f2xl.m_sDestPathXls)
        m_bXlsxExists = bFileExists(m_f2xl.m_sDestPathXlsx)

        m_bDelWorkBookOnClose = My.Settings.DeleteFileOnClose

        ' 20/05/2017 MinColumnWidth and MaxColumnWidth
        ' 28/04/2017 .bRemoveNULL = My.Settings.RemoveNULL
        Dim prm As New clsPrm With {
            .sFieldDelimiters = My.Settings.FieldDelimiters,
            .sDefaultDelimiter = My.Settings.DefaultDelimiter,
            .bUseXls = My.Settings.UseXls,
            .bUseXlsx = My.Settings.UseXlsx,
            .iNbFrozenColumns = My.Settings.NbFrozenColumns,
            .iNbLinesAnalyzed = My.Settings.NbLinesAnalyzed,
            .bPreferMultipleDelimiter = Not bSingleDelimiter,
            .bAutosizeColumns = My.Settings.AutosizeColumns,
            .iMinColumnWidth = My.Settings.MinColumnWidth,
            .iMaxColumnWidth = My.Settings.MaxColumnWidth,
            .bRemoveNULL = My.Settings.RemoveNULL,
            .bLogFile = My.Settings.LogFile
        }
        'If prm.bLogFile Then
        '    m_f2xl.m_sb.AppendLine("Arguments: " & Microsoft.VisualBasic.Interaction.Command)
        '    m_f2xl.m_sb.AppendLine("Path: " & sPath)
        '    If sArgument.Length > 0 Then m_f2xl.m_sb.AppendLine("Argument: " & sArgument)
        'End If

        If Not prm.bUseXls AndAlso Not prm.bUseXlsx Then
            If bDebug Then Stop
            prm.bUseXlsx = True
        End If

        prm.bCreateStandardSheet = My.Settings.CreateStandardSheet

        ShowMessage("Converting...")

        Dim dTimeStart = Now()
        If m_f2xl.bRead(prm, sPath, m_delegMsg) Then
            Dim sDestPath$ = m_f2xl.m_sDestPathXls
            If m_f2xl.m_bXlsx Then sDestPath = m_f2xl.m_sDestPathXlsx

            If My.Settings.LogFile Then ' 20/05/2017
                Dim ci = Globalization.CultureInfo.CurrentCulture()
                Dim dTimeEnd = Now()
                Dim ts = dTimeEnd - dTimeStart
                Const sDateTimeFormat = "dd\/MM\/yyyy HH:mm:ss"
                Dim sTime$ = dTimeStart.ToString(sDateTimeFormat, ci) & " -> " &
                    dTimeEnd.ToString(sDateTimeFormat, ci) & " : " & sDisplayTime(ts.TotalSeconds)
                Dim sb As New StringBuilder()
                sb.AppendLine()
                sb.AppendLine(sTime)
                sb.AppendLine("  -> " & sPath)
                sb.Append(m_f2xl.m_sb)
                Dim sLogPath$ = Application.StartupPath & "\File2XL.log"
                bWriteFile(sLogPath, sb, bAppend:=True)
            End If

            If Not bLetOpenFile(sDestPath) Then m_bDelWorkBookOnClose = False
        End If

        Return True

    End Function

    Private Sub Quit()

        If m_bClosing Then Exit Sub
        m_bClosing = True

        If m_delegMsg.m_bCancel Then GoTo QuitNow
        If m_delegMsg.m_bPause Then m_delegMsg.m_bCancel = True : GoTo QuitNow

        Dim sPath2$ = m_f2xl.m_sDestPathXlsx
        Dim sPath$ = m_f2xl.m_sDestPathXls
        Dim bWorkBookExists = m_bXlsExists
        If m_f2xl.m_bXlsx Then
            sPath = m_f2xl.m_sDestPathXlsx : bWorkBookExists = m_bXlsxExists
            sPath2 = m_f2xl.m_sDestPathXls ' Delete second path too, if necessary
        End If

        If String.IsNullOrEmpty(sPath) Then GoTo QuitNow

        If Not bWorkBookExists AndAlso m_bDelWorkBookOnClose Then

            ' Wait, quit Excel and delete workbook

            If bRelease Then Me.WindowState = FormWindowState.Minimized

            Me.cmdShow.Enabled = False
            Me.cmdStart.Enabled = False
            Me.cmdCancel.Enabled = True
            m_delegMsg.m_bCancel = False

            m_f2xl.FreeMemory()
            ShowMessage("Freeing memory...")
            FreeDotNetRAM()
            ShowMessage("Done.")

            While bFileIsAvailable(sPath, bNonExistentOk:=True, bCheckForSlowRead:=True)
                ShowMessage("Waiting for the workbook to be open...")
                If m_delegMsg.m_bCancel Then Exit While
                Threading.Thread.Sleep(500)
            End While

            While Not bFileIsAvailable(sPath, bNonExistentOk:=True, bCheckForSlowRead:=True)
                ShowMessage("Waiting for the workbook to be closed, and for deleting it...")
                If m_delegMsg.m_bCancel Then Exit While
                Threading.Thread.Sleep(500)
            End While

            If Not m_delegMsg.m_bCancel Then

                If My.Settings.DeleteFileConfirm Then
                    ShowMessage("Confirm the deletion of the workbook...")
                    Me.WindowState = FormWindowState.Normal
                    If MsgBoxResult.Cancel = MsgBox(
                        "Delete temporary workbook ? " &
                        IO.Path.GetFileName(sPath) & vbLf & sPath,
                        MsgBoxStyle.Question Or MsgBoxStyle.OkCancel, m_sMsgTitle) Then GoTo QuitNow
                End If

                If Not bDeleteFile(sPath) Then
                    If bDebug Then Stop
                End If

                ' If necessary delete the other file
                If Not m_bXlsxExists AndAlso Not m_bXlsExists Then
                    If Not bDeleteFile(sPath2) Then
                        If bDebug Then Stop
                    End If
                End If

            End If

        Else

            ' Wait and quit Excel
            If bRelease Then
                While Not bFileIsAvailable(sPath, bNonExistentOk:=True, bCheckForSlowRead:=True)
                    ShowMessage("Waiting for the workbook to be closed...")
                    If m_delegMsg.m_bCancel Then Exit While
                    Threading.Thread.Sleep(500)
                End While
            End If

        End If

QuitNow:
        Me.Close()

    End Sub

    Private Sub cmdStart_Click(sender As Object, e As EventArgs) Handles cmdStart.Click

        m_delegMsg.m_bPause = Not m_delegMsg.m_bPause
        Me.cmdCancel.Enabled = True

        If m_delegMsg.m_bPause Then
            Me.cmdStart.Text = "Continue"
            Me.cmdShow.Enabled = True
            WaitCursor(bDisable:=True)
        Else
            Me.cmdStart.Text = "Pause"
            Me.cmdShow.Enabled = False
            WaitCursor()
        End If

        Application.DoEvents()

    End Sub

    Private Sub cmdShow_Click(sender As Object, e As EventArgs) Handles cmdShow.Click

        Dim sPath$ = m_f2xl.m_sDestPathXls
        If m_f2xl.m_bXlsx Then sPath = m_f2xl.m_sDestPathXlsx
        If m_f2xl.bWrite() Then bLetOpenFile(sPath)
        m_delegMsg.m_bCancel = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        m_delegMsg.m_bCancel = True
    End Sub

    Private Sub ShowButtons()
        Me.cmdCancel.Visible = True
        Me.cmdStart.Visible = True
        Me.cmdShow.Visible = True
    End Sub

    Private Sub Activation(bActivate As Boolean)

        Me.cmdCancel.Enabled = Not bActivate
        Me.cmdStart.Enabled = Not bActivate
        Me.cmdShow.Enabled = bActivate
        If Not bActivate Then
            WaitCursor()
            Me.cmdStart.Text = "Pause"
        Else
            WaitCursor(bDisable:=True)
            Me.cmdStart.Text = "Start"
        End If

        Application.DoEvents()

    End Sub

    Private Sub HideContextMenus()
        Me.lblContextMenu.Visible = False
        Me.cmdAddContextMenu.Visible = False
        Me.cmdRemoveContextMenu.Visible = False
        Me.cmdCreateTestFiles.Visible = False
    End Sub

    Private Sub ShowLongMessage(sMsg$)
        Me.lblInfo.Text = sMsg
        Application.DoEvents() ' Required
    End Sub

    Private Sub ShowMessage(sMsg$)

        Me.ToolStripLabel1.Text = sMsg
        If Me.WindowState <> FormWindowState.Minimized Then
            TruncateChildTextAccordingToControlWidth(Me.ToolStripLabel1, Me, appendEllipsis:=True)
            Dim iLong% = Me.ToolStripLabel1.Text.Length
            If iLong < 30 AndAlso iLong < sMsg.Length And bDebug Then
                Debug.WriteLine(sMsg & " -> ")
                Debug.WriteLine(Me.ToolStripLabel1.Text)
                Stop
            End If
        End If
        Application.DoEvents() ' Required

    End Sub

    Private Sub ShowMessageDeleg(sender As Object, e As clsMsgEventArgs) Handles m_delegMsg.EvShowMessage
        Me.ShowMessage(e.sMessage)
    End Sub

    Private Sub ShowLongMessageDeleg(sender As Object, e As clsMsgEventArgs) _
        Handles m_delegMsg.EvShowLongMessage
        Me.ShowLongMessage(e.sMessage)
    End Sub

    Private Sub SetWaitCursor(sender As Object, e As clsWaitCursorEventArgs) _
        Handles m_delegMsg.EvWaitCursor
        WaitCursor(e.bDisable)
    End Sub

    Private Shared Sub WaitCursor(Optional bDisable As Boolean = False)

        If bDisable Then
            Application.UseWaitCursor = False
        Else
            Application.UseWaitCursor = True
        End If

    End Sub

    Private Class clsTest
        Public Const sTestHeader$ = "TestHeader"
        Public Const sTest255Col$ = "Test255Col"
        Public Const sTest256Col$ = "Test256Col"
        Public Const sTest257Col$ = "Test257Col"
        Public Const sTest16384Col$ = "Test16384Col"
        Public Const sTest16385Col$ = "Test16385Col"
        Public Const sTest65536Lines$ = "Test65536Lines"
        Public Const sTest65536LinesBig$ = "Test65536LinesBig"
        Public Const sTest65537Lines$ = "Test65537Lines"
        Public Const sTest1048576Lines$ = "Test1048576Lines"
        Public Const sTest1048577Lines$ = "Test1048577Lines"
        Public Const sTestMaxCarCell32767$ = "TestMaxCarCell32767"
        Public Const sTestMaxCarCell32768$ = "TestMaxCarCell32768"
        Public Const sTestBigExcel2003$ = "TestBigExcel2003"
        Public Const sTestVeryBigExcel2003$ = "TestVeryBigExcel2003"
        Public Const sTestBigExcel2007$ = "TestBigExcel2007"
    End Class

    Private Sub cmdCreateTestFiles_Click(sender As Object, e As EventArgs) Handles cmdCreateTestFiles.Click

        Me.cmdCreateTestFiles.Enabled = False

        CreateTestFile(clsTest.sTestHeader)
        CreateTestFile(clsTest.sTest255Col)
        CreateTestFile(clsTest.sTest256Col)
        CreateTestFile(clsTest.sTest257Col)
        CreateTestFile(clsTest.sTest16384Col)
        CreateTestFile(clsTest.sTest16385Col)
        CreateTestFile(clsTest.sTest65536Lines)
        CreateTestFile(clsTest.sTest65536LinesBig)
        CreateTestFile(clsTest.sTest65537Lines)
        CreateTestFile(clsTest.sTest1048576Lines)
        CreateTestFile(clsTest.sTest1048577Lines)
        CreateTestFile(clsTest.sTestMaxCarCell32767)
        CreateTestFile(clsTest.sTestMaxCarCell32768)
        CreateTestFile(clsTest.sTestBigExcel2003)
        CreateTestFile(clsTest.sTestVeryBigExcel2003)
        CreateTestFile(clsTest.sTestBigExcel2007)

        'EndTest:
        Me.cmdCreateTestFiles.Enabled = True

        MsgBox("OK !")

    End Sub

    Private Shared Sub CreateTestFile(sTestFile$)

        Dim bTestHeader As Boolean = False
        Dim bTestMaxTxtCell As Boolean = False
        Dim iNbCol%, iNbLines%, iNbCarMax%
        iNbCol = 10
        iNbLines = 10
        iNbCarMax = clsFile2XL.iNbCarMaxCell
        Select Case sTestFile
            Case clsTest.sTestHeader : bTestHeader = True
            Case clsTest.sTest255Col : iNbCol = 255
            Case clsTest.sTest256Col : iNbCol = 256
            Case clsTest.sTest257Col : iNbCol = 257
            Case clsTest.sTest16384Col : iNbCol = 16384 : iNbLines = 2
            Case clsTest.sTest16385Col : iNbCol = 16385 : iNbLines = 2
            Case clsTest.sTest65536Lines : iNbLines = 65536 : iNbCol = 2
            Case clsTest.sTest65536LinesBig : iNbLines = 65536 : iNbCol = 10
            Case clsTest.sTest65537Lines : iNbLines = 65537 : iNbCol = 2
            Case clsTest.sTest1048576Lines : iNbLines = 1048576 : iNbCol = 2
            Case clsTest.sTest1048577Lines : iNbLines = 1048577 : iNbCol = 2
            Case clsTest.sTestMaxCarCell32767 : bTestMaxTxtCell = True
            Case clsTest.sTestMaxCarCell32768 : bTestMaxTxtCell = True : iNbCarMax = clsFile2XL.iNbCarMaxCell + 1
            Case clsTest.sTestBigExcel2003 : iNbLines = 10000 : iNbCol = 100
            Case clsTest.sTestVeryBigExcel2003 : iNbLines = 65536 : iNbCol = 256
            Case clsTest.sTestBigExcel2007 : iNbLines = 10000 : iNbCol = 500
        End Select

        Const sDelimiter = ";" 'vbTab

        Dim sb As New StringBuilder
        For i As Integer = 0 To iNbLines - 1
            If bTestHeader Then iNbCol = i
            For j As Integer = 0 To iNbCol - 1

                If bTestMaxTxtCell AndAlso i = 5 AndAlso (j = 5 OrElse j = 7) Then
                    For k As Integer = 0 To iNbCarMax - 1
                        sb.Append("x")
                    Next
                    sb.Append(sDelimiter)
                    Continue For
                End If

                'sb.Append((j + 1 + i * iNbCol))
                sb.Append((j + 1 + i))

                If j < iNbCol - 1 Then sb.Append(sDelimiter)
            Next
            sb.AppendLine()
        Next

        Dim sPath$ = Application.StartupPath & "\" & sTestFile & ".dat"
        If Not bWriteFile(sPath, sb) Then Exit Sub

    End Sub

#Region "Context menus"

    Private Sub CheckContextMenu()

        Dim sKey$ = sContextMenu_FileTypeAll & "\" & sShellKey & "\" & sContextMenu_CmdKeyOpen
        If bClassesRootRegistryKeyExists(sKey) Then
            Me.cmdAddContextMenu.Enabled = False
            Me.cmdRemoveContextMenu.Enabled = True
        Else
            Me.cmdAddContextMenu.Enabled = True
            Me.cmdRemoveContextMenu.Enabled = False
        End If

    End Sub

    Private Sub cmdAddContextMenu_Click(sender As Object, e As EventArgs) _
        Handles cmdAddContextMenu.Click
        AddContextMenus()
        CheckContextMenu()
    End Sub

    Private Sub cmdRemoveContextMenu_Click(sender As Object, e As EventArgs) _
        Handles cmdRemoveContextMenu.Click
        RemoveContextMenus()
        CheckContextMenu()
    End Sub

    Private Shared Sub AddContextMenus()

        If MsgBoxResult.Cancel = MsgBox("Add context menu ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question, m_sMsgTitle) Then Exit Sub

        AddContextMenus(sContextMenu_FileTypeAll)

    End Sub

    Private Shared Sub RemoveContextMenus()

        If MsgBoxResult.Cancel = MsgBox("Remove context menu ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question, m_sMsgTitle) Then Exit Sub

        RemoveContextMenus(sContextMenu_FileTypeAll)

    End Sub

    Private Shared Sub AddContextMenus(sKey$)

        Dim sExePath$ = Application.ExecutablePath
        Const bPrompt As Boolean = False
        Const sPath = """%1"""
        bAddContextMenu(sKey, sContextMenu_CmdKeyOpen,
            bPrompt, , sContextMenu_CmdKeyOpenDescr, sExePath, sPath)
        bAddContextMenu(sKey, sContextMenu_CmdKeyOpen2,
            bPrompt, , sContextMenu_CmdKeyOpen2Descr, sExePath, sPath & " " & sSingleDelimiterArg)

    End Sub

    Private Shared Sub RemoveContextMenus(sKey$)

        bAddContextMenu(sKey, sContextMenu_CmdKeyOpen, bRemove:=True, bPrompt:=False)
        bAddContextMenu(sKey, sContextMenu_CmdKeyOpen2, bRemove:=True, bPrompt:=False)

    End Sub

#End Region

End Class