Option Explicit

'====================================================
' Class: clsCompany
' Purpose: Company data model
' Version: 0.1.0-dev
'====================================================

Private mTicker As String
Private mCompanyName As String
Private mSector As String
Private mIndustry As String


Public Property Get Ticker() As String
    Ticker = mTicker
End Property

Public Property Let Ticker(ByVal Value As String)
    mTicker = Value
End Property


Public Property Get CompanyName() As String
    CompanyName = mCompanyName
End Property

Public Property Let CompanyName(ByVal Value As String)
    mCompanyName = Value
End Property


Public Property Get Sector() As String
    Sector = mSector
End Property

Public Property Let Sector(ByVal Value As String)
    mSector = Value
End Property


Public Property Get Industry() As String
    Industry = mIndustry
End Property

Public Property Let Industry(ByVal Value As String)
    mIndustry = Value
End Property