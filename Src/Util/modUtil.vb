
' File modUtil.vb : Utility module
' ---------------

Imports System.Runtime.CompilerServices ' For MethodImpl(MethodImplOptions.AggressiveInlining)

Module modUtil

    ' This field is rather a static variable than a member variable, it should be named s_sMsgTitle
    <CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1504:ReviewMisleadingFieldNames")>
    Public m_sMsgTitle$ = sMsgTitle

    Public Sub SetMsgTitle(sMsgTitle$)
        m_sMsgTitle = sMsgTitle
    End Sub

    Public Sub ShowErrorMsg(ex As Exception,
        Optional sFunctionTitle$ = "", Optional sInfo$ = "",
        Optional sDetailedErrMsg$ = "",
        Optional bCopyErrMsgClipboard As Boolean = True,
        Optional ByRef sFinalErrMsg$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then Cursor.Current = Cursors.Default

        Dim sMsg$ = ""
        If sFunctionTitle <> "" Then sMsg = "Function : " & sFunctionTitle
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailedErrMsg <> "" Then sMsg &= vbCrLf & sDetailedErrMsg
        If ex.Message <> "" Then
            sMsg &= vbCrLf & ex.Message.Trim
            If Not IsNothing(ex.InnerException) Then _
                sMsg &= vbCrLf & ex.InnerException.Message
        End If
        If bCopyErrMsgClipboard Then
            CopyToClipboard(sMsg)
            sMsg &= vbCrLf & "(this error message has been copied into the clipboard)"
        End If

        sFinalErrMsg = sMsg

        MsgBox(sMsg, MsgBoxStyle.Critical, m_sMsgTitle)

    End Sub

    Public Sub CopyToClipboard(sInfo$)

        ' Copy text into Windows clipboard (until the application is closed)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' The clipboard can be unavailable
            ShowErrorMsg(ex, "CopyToClipboard", bCopyErrMsgClipboard:=False)
        End Try

    End Sub

    Public Function is64BitProcess() As Boolean
        Return (IntPtr.Size = 8)
    End Function

#Region "Documentation"
    ''' <summary>
    ''' If a child ToolStripStatusLabel is wider than it's parent then this method will attempt to
    '''  make the child's text fit inside of the parent's boundaries. An ellipsis can be appended
    '''  at the end of the text to indicate that it has been truncated to fit.
    ''' </summary>
    ''' <param name="child">Child ToolStripStatusLabel</param>
    ''' <param name="parent">Parent control where the ToolStripStatusLabel resides</param>
    ''' <param name="appendEllipsis">Append an "..." to the end of the truncated text</param>
#End Region
    Public Sub TruncateChildTextAccordingToControlWidth(child As ToolStripLabel,
        parent As Control, appendEllipsis As Boolean)

        ' http://stackoverflow.com/questions/5708375/how-can-i-determine-how-much-of-a-string-will-fit-in-a-certain-width

        ' If the child's width is greater than that of the parent's
        Const rPadding = 0.1
        'If child.Size.Width >= parent.Size.Width * 0.9 Then
        If child.Size.Width >= parent.Size.Width * (1 - rPadding) Then

            ' Get the number of times that the child is oversized [child/parent]
            Dim decOverSized As Decimal = CDec(child.Size.Width) / CDec(parent.Size.Width)

            ' Get the new Text length based on the number of times that the child's width is oversized.
            'Dim iNewLength% = CInt(child.Text.Length / (2D * decOverSized))
            Dim iNewLength% = CInt(child.Text.Length / ((1 + rPadding) * decOverSized))

            ' Doubling as a buffer (Magic Number).
            ' If the ellipsis is to be appended
            If appendEllipsis Then
                ' then 3 more characters need to be removed to make room for it.
                iNewLength = iNewLength - 3
            End If

            ' If the new length is negative for whatever reason
            If iNewLength < 0 Then iNewLength = 0 ' Then default it to zero

            ' Truncate the child's Text accordingly
            If child.Text.Length >= iNewLength Then _
                child.Text = child.Text.Substring(0, iNewLength)

            ' If the ellipsis is to be appended
            If appendEllipsis Then child.Text += "..." ' Then do this last

        End If

    End Sub

    ' GC.Collect is rarely usefull
    <CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId:="System.GC.Collect")>
    Public Sub FreeDotNetRAM(Optional bComResources As Boolean = False)

        ' Clean up managed, and unmanaged COM resources if bComResources is True

        ' Clean up the unmanaged COM resources by forcing a garbage
        ' collection as soon as the calling function is off the stack
        ' (at which point these objects are no longer rooted).
        GC.Collect()

        If bComResources Then
            GC.WaitForPendingFinalizers()
            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually is finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End If

    End Sub

    Public Function sRAMInfo$(Optional sMsg$ = "RAM : ")

        Dim ci = Globalization.CultureInfo.CurrentCulture()
        Dim x As Process = System.Diagnostics.Process.GetCurrentProcess
        Dim lAllocatedRamByApp& = x.WorkingSet64

        Dim sAllocatedRamByApp$ = sDisplaySizeInBytes(lAllocatedRamByApp)

        ' In 32 bits, only 1.6 Gb can be allocated (inside Visual Studio or not)

        If Not is64BitProcess() Then
            Dim lRamAvailable32 As ULong = CULng(1.6 * 1024 * 1024 * 1024) ' 1.6 Go
            If lRamAvailable32 < My.Computer.Info.AvailablePhysicalMemory Then
                lRamAvailable32 = My.Computer.Info.AvailablePhysicalMemory
            End If
            Dim sRamAvailable32$ = sDisplaySizeInBytes(CLng(lRamAvailable32))
            Dim rPCRAMUsed32! = CSng(lAllocatedRamByApp / lRamAvailable32)
            Dim sRam32$ = sMsg & sAllocatedRamByApp & " / " & sRamAvailable32 &
                " (" & rPCRAMUsed32.ToString("0.0 %", ci) & ")"
            Return sRam32
        End If

        Dim lRamAvailable As ULong = My.Computer.Info.AvailablePhysicalMemory
        'Dim sRamAvailable$ = sDisplaySizeInBytes(CLng(lRamAvailable))
        Dim lRamTot As ULong = My.Computer.Info.TotalPhysicalMemory
        Dim sRamTot$ = sDisplaySizeInBytes(CLng(lRamTot))
        Dim lTotAllocated As ULong = lRamTot - lRamAvailable
        Dim sTotAllocatedRAM$ = sDisplaySizeInBytes(CLng(lTotAllocated))
        Dim lAllocatedByOtherProc As ULong = CULng(lTotAllocated - lAllocatedRamByApp)
        Dim sAllocatedByOtherProc$ = sDisplaySizeInBytes(CLng(lAllocatedByOtherProc))

        Dim rPCRAMUsed! = CSng(lTotAllocated / lRamTot)
        Dim sRam$ = sMsg & sAllocatedRamByApp & " + " & sAllocatedByOtherProc & " = " &
            sTotAllocatedRAM & " / " & sRamTot & " (" & rPCRAMUsed.ToString("0.0 %", ci) & ")"
        Return sRam

    End Function

    <MethodImpl(MethodImplOptions.AggressiveInlining)>
    Public Function rFastConv#(sValue$, Optional rDef! = 0.0!, Optional ByRef bOK As Boolean = True)

        bOK = False
        If String.IsNullOrEmpty(sValue) Then Return rDef

        Dim rVal#

        If Double.TryParse(sValue, rVal) Then
            bOK = True : Return rVal
        Else

            Dim sVal2$ = sValue.Replace(sDot, sComma)
            If Double.TryParse(sVal2, rVal) Then bOK = True : Return rVal
            Dim sVal3$ = sValue.Replace(sComma, sDot)
            If Double.TryParse(sVal3, rVal) Then bOK = True : Return rVal

            Return rDef
        End If

    End Function

End Module