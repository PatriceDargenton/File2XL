
' File modConst.vb
' ----------------

Module _modConst

Public Const sAppDate$ = "22/10/2021" ' 1.05:25/01/2019 1.04:05/01/2018 1.03:20/05/2017 1.02:08/05/2017 1.01:16/10/2016

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