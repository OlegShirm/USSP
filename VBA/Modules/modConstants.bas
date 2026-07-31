Attribute VB_Name = "modConstants"
Option Explicit

'=========================================================
' USSP - US Stock Screener Professional
'
' Module:
'   modConstants
'
' Version:
'   0.1.0-dev
'
' Purpose:
'   Global application constants
'
' Milestone:
'   M2.003 Constants Framework
'
'=========================================================

'---------------------------------------------------------
' Application
'---------------------------------------------------------

Public Const APP_NAME As String = "USSP"
Public Const APP_FULLNAME As String = "US Stock Screener Professional"
Public Const APP_VERSION As String = "0.1.0-dev"

'---------------------------------------------------------
' Worksheet names
'---------------------------------------------------------

Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_SCREENER As String = "Stock Screener"
Public Const WS_COMPANYDATA As String = "Company Data"
Public Const WS_MARKETDATA As String = "Market Data"
Public Const WS_WATCHLIST As String = "Watch List"
Public Const WS_CANDIDATES As String = "Candidate List"
Public Const WS_PRESETS As String = "Scan Presets"
Public Const WS_RULES As String = "Rules Engine"
Public Const WS_PROVIDERS As String = "Data Providers"
Public Const WS_SETTINGS As String = "Settings"
Public Const WS_LOGS As String = "Logs"

'---------------------------------------------------------
' General
'---------------------------------------------------------

Public Const EMPTY_STRING As String = ""

Public Const DATE_FORMAT As String = "yyyy-mm-dd"

Public Const PERCENT_FORMAT As String = "0.00%"

Public Const CURRENCY_FORMAT As String = "$#,##0.00"

'---------------------------------------------------------
' Scan limits
'---------------------------------------------------------

Public Const MAX_COMPANIES As Long = 10000

Public Const MAX_WATCHLIST As Long = 500

'---------------------------------------------------------
' Status values
'---------------------------------------------------------

Public Const STATUS_READY As String = "Ready"

Public Const STATUS_RUNNING As String = "Running"

Public Const STATUS_FINISHED As String = "Finished"

Public Const STATUS_ERROR As String = "Error"
