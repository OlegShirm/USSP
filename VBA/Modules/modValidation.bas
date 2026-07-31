Attribute VB_Name = "modValidation"
Option Explicit

'=========================================================
' USSP - Validation Module
'
' Module:
'   modValidation
'
' Purpose:
'   Common validation functions
'
'=========================================================

Public Function IsValidTicker(ByVal Ticker As String) As Boolean

    Ticker = Trim$(Ticker)

    If Len(Ticker) = 0 Then Exit Function

    If Len(Ticker) > 10 Then Exit Function

    IsValidTicker = True

End Function


Public Function IsPositiveNumber(ByVal Value As Double) As Boolean

    IsPositiveNumber = (Value > 0)

End Function


Public Function IsNonNegativeNumber(ByVal Value As Double) As Boolean

    IsNonNegativeNumber = (Value >= 0)

End Function


Public Function IsValidPercentage(ByVal Value As Double) As Boolean

    IsValidPercentage = _
        (Value >= 0# And Value <= 100#)

End Function


Public Function IsValidDateValue(ByVal Value As Variant) As Boolean

    IsValidDateValue = IsDate(Value)

End Function


Public Function IsNotEmpty(ByVal Text As String) As Boolean

    IsNotEmpty = (Len(Trim$(Text)) > 0)

End Function

