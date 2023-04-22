
' File modUtilReg.vb : Registry module utility
' ------------------

Imports Microsoft.Win32

Module modUtilReg

    ' Microsoft Win32 to Microsoft .NET Framework API Map : Registry Functions
    ' http://msdn.microsoft.com/en-us/library/aa302340.aspx#win32map_registryfunctions

    Public Const sShellKey$ = "shell"
    Public Const sCmdKey$ = "command"

    Private Const sMsgErrPossibleCause$ =
        "Possible cause : adding context menus requires administrator privileges" & vbCrLf &
        "Run as admin. the application for this operation"

    Public Function bAddContextMenuFileType(sExtension$, sFileType$,
        Optional sExtensionDescription$ = "",
        Optional bRemove As Boolean = False) As Boolean

        ' Add (or Remove) in the registry a ClassesRoot file type
        '  to associate a file extension to a default application
        ' (via double-click or for example the context menu Open, see the next function bAddContextMenu)
        ' Example : associate .dat file extension to MyApplication.exe

        Try

            If bRemove Then
                If bClassesRootRegistryKeyExists(sExtension) Then
                    Registry.ClassesRoot.DeleteSubKeyTree(sExtension)
                End If
            Else
                If Not bClassesRootRegistryKeyExists(sExtension) Then
                    Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sExtension)
                        rk.SetValue("", sFileType)
                        If sExtensionDescription.Length > 0 Then
                            rk.SetValue("Content Type", sExtensionDescription)
                        End If
                    End Using
                End If
            End If
            Return True

        Catch ex As Exception
            ShowErrorMsg(ex, "bAddContextMenuFileType", sMsgErrPossibleCause)
            Return False
        End Try

    End Function

    Public Function bAddContextMenu(sFileType$, sCmd$,
        Optional bPrompt As Boolean = True,
        Optional bRemove As Boolean = False,
        Optional sCmdDescription$ = "",
        Optional sExePath$ = "",
        Optional sCmdDef$ = """%1""",
        Optional sFileTypeDescription$ = "",
        Optional bRemoveFileType As Boolean = False) As Boolean

        ' Add (or Remove) in the registry a context menu for a ClassesRoot file type
        ' (see the previous function bAddContextMenuFileType)
        '  to associate a command menu for a file extension (or for every file) to a default application
        ' (via double-click or the context menu Open, for example, in the Windows File Explorer)
        ' Example one   : associate the menu Open   for .dat  file extension to MyApplication.exe
        ' Example two   : associate the menu Open   for every file           to MyApplication.exe
        ' Example three : associate the menu Print  for .doc  file extension to MyApplication.exe
        ' Example four  : associate the menu Search for every folder         to MyApplication.exe

        Try

            ' Fisrt check the main key
            If Not bClassesRootRegistryKeyExists(sFileType) Then
                If bRemove Then bAddContextMenu = True : Exit Function
                Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sFileType)
                    If sFileTypeDescription.Length > 0 Then
                        rk.SetValue("", sFileTypeDescription)
                    End If
                End Using
            End If

            Dim sCleDescriptionCmd$ = sFileType & "\" & sShellKey & "\" & sCmd

            If bRemove Then

                If bRemoveFileType Then
                    If bClassesRootRegistryKeyExists(sFileType) Then
                        Registry.ClassesRoot.DeleteSubKeyTree(sFileType)
                        If bPrompt Then _
                            MsgBox("The context menu [" & sFileType & "]" & vbLf &
                                "has been successfully removed from registry",
                                MsgBoxStyle.Information, m_sMsgTitle)
                    Else
                        If bPrompt Then _
                            MsgBox("The context menu [" & sFileType & "]" & vbLf &
                                "can't be found in the registry",
                                MsgBoxStyle.Information, m_sMsgTitle)
                    End If
                Else

                    If bClassesRootRegistryKeyExists(sCleDescriptionCmd) Then
                        Registry.ClassesRoot.DeleteSubKeyTree(sCleDescriptionCmd)
                        If bPrompt Then _
                            MsgBox("The context menu [" & sCmdDescription & "]" & vbLf &
                                "has been successfully removed from registry for the files of the type :" & vbLf &
                                "[" & sFileType & "]",
                                MsgBoxStyle.Information, m_sMsgTitle)
                    Else
                        If bPrompt Then _
                            MsgBox("The context menu [" & sCmdDescription & "]" & vbLf &
                                "can't be found in the registry for the files of the type :" & vbLf &
                                "[" & sFileType & "]",
                                MsgBoxStyle.Information, m_sMsgTitle)
                    End If

                End If
                bAddContextMenu = True
                Exit Function
            End If

            Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sCleDescriptionCmd)
                rk.SetValue("", sCmdDescription)
            End Using 'rk.Close()

            Dim sCleCmd$ = sFileType & "\" & sShellKey & "\" & sCmd & "\" & sCmdKey
            Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sCleCmd)
                ' Add quotes " if the path contains spaces
                If sExePath.IndexOf(" ", StringComparison.Ordinal) > -1 Then _
                    sExePath = """" & sExePath & """"
                rk.SetValue("", sExePath & " " & sCmdDef)
            End Using

            If bPrompt Then _
                MsgBox("The context menu [" & sCmdDescription & "]" & vbLf &
                    "has been successfully added from registry for the files of the type :" & vbLf &
                    "[" & sFileType & "]", MsgBoxStyle.Information, m_sMsgTitle)

            Return True

        Catch ex As Exception
            ShowErrorMsg(ex, "bAddContextMenu", sMsgErrPossibleCause)
            Return False
        End Try

    End Function

    Public Function bClassesRootRegistryKeyExists(sKey$, Optional sSubKey$ = "") As Boolean

        Try
            Using rkCRCle As RegistryKey = Registry.ClassesRoot.OpenSubKey(sKey & "\" & sSubKey)
                If IsNothing(rkCRCle) Then Return False
            End Using
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function bClassesRootRegistryKeyExists(sKey$, sSubKey$, ByRef sSubKeyValue$) As Boolean

        sSubKeyValue = ""
        Try
            Using rkCRCle As RegistryKey = Registry.ClassesRoot.OpenSubKey(sKey)
                If IsNothing(rkCRCle) Then Return False
                Dim oValue As Object = rkCRCle.GetValue(sSubKey)
                If IsNothing(oValue) Then Return False
                Dim sSubKeyValue0$ = CStr(oValue)
                If IsNothing(sSubKeyValue0) Then Return False
                sSubKeyValue = sSubKeyValue0
            End Using
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function bLocalMachineRegistryKeyExists(sKey$, Optional sSubKey$ = "",
        Optional ByRef sSubKeyValue$ = "", Optional sNewSubKeyValue$ = "") As Boolean

        sSubKeyValue = ""
        Try
            Dim bWrite As Boolean = False
            If sNewSubKeyValue.Length > 0 Then bWrite = True
            Using rkLMCle As RegistryKey = Registry.LocalMachine.OpenSubKey(sKey,
                writable:=bWrite)
                Dim oValue As Object = rkLMCle.GetValue(sSubKey)
                If IsNothing(oValue) Then Return False
                Dim sSubKeyVal0$ = CStr(oValue)
                If IsNothing(sSubKeyVal0) Then Return False
                sSubKeyValue = sSubKeyVal0
                If bWrite Then
                    oValue = CInt(sNewSubKeyValue)
                    rkLMCle.SetValue(sSubKey, oValue, RegistryValueKind.DWord)
                End If
            End Using
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function bCurrentUserRegistryKeyExists(sKey$, Optional sSubKey$ = "",
        Optional ByRef sSubKeyValue$ = "") As Boolean

        sSubKeyValue = ""
        Try
            Using rkCUCle As RegistryKey = Registry.CurrentUser.OpenSubKey(sKey)
                Dim oValue As Object = rkCUCle.GetValue(sSubKey)
                If IsNothing(oValue) Then Return False
                Dim sSubKeyValue0$ = CStr(oValue)
                If IsNothing(sSubKeyValue0) Then Return False
                sSubKeyValue = sSubKeyValue0
            End Using
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function asCurrentUserRegistrySubKeys(sKey$) As String()

        Try
            Using rkCUCle As RegistryKey = Registry.CurrentUser.OpenSubKey(sKey)
                If IsNothing(rkCUCle) Then Return Nothing
                Return rkCUCle.GetSubKeyNames
            End Using
        Catch
            Return Nothing
        End Try

    End Function

End Module