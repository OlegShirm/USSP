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
Private mFounded As Date
Private mIPODate As Date


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


Public Property Get Founded() As Date
    Founded = mFounded
End Property

Public Property Let Founded(ByVal Value As Date)
    mFounded = Value
End Property


Public Property Get IPODate() As Date
    IPODate = mIPODate
End Property

Public Property Let IPODate(ByVal Value As Date)
    mIPODate = Value
End Property

Public Property Get CompanyAge() As Integer

    If mFounded = 0 Then
        CompanyAge = 0
        Exit Property
    End If

    CompanyAge = DateDiff("yyyy", mFounded, Date)

End Property


Public Property Get IPOAge() As Integer

    If mIPODate = 0 Then
        IPOAge = 0
        Exit Property
    End If

    IPOAge = DateDiff("yyyy", mIPODate, Date)

End Property