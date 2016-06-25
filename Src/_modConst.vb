
' File modConst.vb
' ----------------

Module _modConst

Public Const sAppDate$ = "25/06/2016"

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

Public ReadOnly sAppVersion$ = _
    My.Application.Info.Version.Major & "." & _
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

Public Const sTxtSheet$ = "Text sheet"
Public Const sStdrSheet$ = "Standard sheet"

Public Const sPostFixWithQuotes$ = "WithQuotes"

End Module

Public Class clsFieldType
    Public Const sNumericC2P$ = "NumericC2P"
    Public Const sNumericP2C$ = "NumericP2C" ' Period to Comma
    Public Const sNumeric$ = "Numeric"
    Public Const sNumericWithQuotes$ = "NumericWithQuotes"
    Public Const sNumericC2PWithQuotes$ = "NumericC2PWithQuotes"
    Public Const sNumericP2CWithQuotes$ = "NumericP2CWithQuotes"
    Public Const sText$ = "Text"
    Public Const sTextWithQuotes$ = "TextWithQuotes"
End Class

Public Class clsField
    Public sField$, sType$
    Public iNumField%, iNbOcc%
    Public bCanEndWithMinus As Boolean = False ' Numeric followed by -
    Public Sub New(iNumField0%, sField0$, sType0$)
        iNumField = iNumField0
        sField = sField0
        sType = sType0
        iNbOcc = 1
    End Sub
End Class