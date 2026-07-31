Attribute VB_Name = "modGlobals"
Option Explicit

'=========================================================
' USSP - US Stock Screener Professional
'
' Module:
'   modGlobals
'
' Version:
'   0.1.0-dev
'
' Purpose:
'   Global application objects and state
'
' Milestone:
'   M2.004 Global State
'
'=========================================================

'---------------------------------------------------------
' Global data
'---------------------------------------------------------

Public gCompanies As clsCompanyCollection

'---------------------------------------------------------
' Current state
'---------------------------------------------------------

Public gCurrentTicker As String

Public gCurrentProvider As String

Public gApplicationInitialized As Boolean

Public gWorkbookVersion As String
