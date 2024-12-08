
' File modConst.vb
' ----------------

Module _modConst

    Public Const sAppDate$ = "08/12/2024"

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

    Public ReadOnly sAppVersion$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor & My.Application.Info.Version.Build

    Public ReadOnly sAppName$ = My.Application.Info.Title
    Public ReadOnly sMsgTitle$ = sAppName

    Public Const iDisplayRate% = 1000
    Public Const sMsgDone$ = "Done."

    Public Const sQuotes$ = """" ' Chr$(34) = "
    Public Const sQuotesCommaQuotesDelimiter$ = sQuotes & sComma & sQuotes
    Public Const sQuotesSemiColonQuotesDelimiter$ = sQuotes & ";" & sQuotes

    Public Const sDot$ = "."
    Public Const sPeriod$ = sDot
    Public Const sComma$ = ","
    Public Const sEmpty$ = ""
    Public Const sNULL$ = "NULL" ' PhpMyAdmin null value in csv export

    Public Const sTxtSheet$ = "Text sheet"
    Public Const sStdrSheet$ = "Standard sheet"

    Public Const sPostFixWithQuotes$ = "WithQuotes"

End Module