VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompanyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Sub Add(ByVal Company As clsCompany)
    mItems.Add Company
End Sub

Public Function Item(ByVal Index As Long) As clsCompany
    Set Item = mItems(Index)
End Function

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Sub Clear()
    Set mItems = New Collection
End Sub
