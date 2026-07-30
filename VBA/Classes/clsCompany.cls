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

Private mMarketCap As Double
Private mCurrentPrice As Double

Private mWeek52High As Double
Private mWeek52Low As Double

Private mDividendYield As Double

Private mPE As Double
Private mEPS As Double
Private mROE As Double
Private mDebtToEquity As Double

Private mWeek52High As Double
Private mWeek52Low As Double

Private mDividendYield As Double

Private mPE As Double
Private mEPS As Double
Private mROE As Double
Private mDebtToEquity As Double

Private mPERatio As Double
Private mPBRatio As Double
Private mPSRatio As Double
Private mPEGRatio As Double

Private mDividendYield As Double
Private mDividendRate As Double

Private mRevenue As Double
Private mNetIncome As Double
Private mEPS As Double
Private mEBITDA As Double
Private mFreeCashFlow As Double

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

Public Property Get MarketCap() As Double
    MarketCap = mMarketCap
End Property

Public Property Let MarketCap(ByVal Value As Double)
    mMarketCap = Value
End Property


Public Property Get CurrentPrice() As Double
    CurrentPrice = mCurrentPrice
End Property

Public Property Let CurrentPrice(ByVal Value As Double)
    mCurrentPrice = Value
End Property

Public Property Get Week52High() As Double
    Week52High = mWeek52High
End Property

Public Property Let Week52High(ByVal Value As Double)
    mWeek52High = Value
End Property


Public Property Get Week52Low() As Double
    Week52Low = mWeek52Low
End Property

Public Property Let Week52Low(ByVal Value As Double)
    mWeek52Low = Value
End Property

Public Property Get DividendYield() As Double
    DividendYield = mDividendYield
End Property

Public Property Let DividendYield(ByVal Value As Double)
    mDividendYield = Value
End Property

Public Property Get PE() As Double
    PE = mPE
End Property

Public Property Let PE(ByVal Value As Double)
    mPE = Value
End Property

Public Property Get EPS() As Double
    EPS = mEPS
End Property

Public Property Let EPS(ByVal Value As Double)
    mEPS = Value
End Property

Public Property Get ROE() As Double
    ROE = mROE
End Property

Public Property Let ROE(ByVal Value As Double)
    mROE = Value
End Property

Public Property Get DebtToEquity() As Double
    DebtToEquity = mDebtToEquity
End Property

Public Property Let DebtToEquity(ByVal Value As Double)
    mDebtToEquity = Value
End Property

Public Property Get PERatio() As Double
    PERatio = mPERatio
End Property

Public Property Let PERatio(ByVal Value As Double)
    mPERatio = Value
End Property


Public Property Get PBRatio() As Double
    PBRatio = mPBRatio
End Property

Public Property Let PBRatio(ByVal Value As Double)
    mPBRatio = Value
End Property


Public Property Get PSRatio() As Double
    PSRatio = mPSRatio
End Property

Public Property Let PSRatio(ByVal Value As Double)
    mPSRatio = Value
End Property


Public Property Get PEGRatio() As Double
    PEGRatio = mPEGRatio
End Property

Public Property Let PEGRatio(ByVal Value As Double)
    mPEGRatio = Value
End Property

Public Property Get DividendYield() As Double
    DividendYield = mDividendYield
End Property

Public Property Let DividendYield(ByVal Value As Double)
    mDividendYield = Value
End Property

Public Property Get DividendRate() As Double
    DividendRate = mDividendRate
End Property

Public Property Let DividendRate(ByVal Value As Double)
    mDividendRate = Value
End Property

Public Property Get Revenue() As Double
    Revenue = mRevenue
End Property

Public Property Let Revenue(ByVal Value As Double)
    mRevenue = Value
End Property


Public Property Get NetIncome() As Double
    NetIncome = mNetIncome
End Property

Public Property Let NetIncome(ByVal Value As Double)
    mNetIncome = Value
End Property


Public Property Get EPS() As Double
    EPS = mEPS
End Property

Public Property Let EPS(ByVal Value As Double)
    mEPS = Value
End Property


Public Property Get EBITDA() As Double
    EBITDA = mEBITDA
End Property

Public Property Let EBITDA(ByVal Value As Double)
    mEBITDA = Value
End Property


Public Property Get FreeCashFlow() As Double
    FreeCashFlow = mFreeCashFlow
End Property

Public Property Let FreeCashFlow(ByVal Value As Double)
    mFreeCashFlow = Value
End Property