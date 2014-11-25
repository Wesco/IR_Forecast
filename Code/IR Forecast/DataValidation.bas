Attribute VB_Name = "DataValidation"
Option Explicit

Public Enum Fcst
    Campbellsville
    DLC
    Unicov
    MoxBB
    Discrete
    Wujiang
End Enum

Sub ValidateForecast(Forecast As Fcst)
    Select Case Forecast
        Case Fcst.Campbellsville
            Campbellsville
        Case Fcst.Discrete
            Discrete
        Case Fcst.DLC
            DLC
        Case Fcst.MoxBB
            MoxBB
        Case Fcst.Unicov
            Unicov
        Case Fcst.Wujiang
            Wujiang
        Case Else
            'This should never happen
            Err.Raise 50000, "ValidateForecast", "Unknown forecast"
    End Select
End Sub

Private Sub Campbellsville()
    Sheets("Campbellsville").Select
End Sub

Private Sub DLC()
    Sheets("DLC").Select
End Sub

Private Sub Unicov()
    Sheets("Unicov").Select
End Sub

Private Sub MoxBB()
    Sheets("MoxBB").Select
End Sub

Private Sub Discrete()
    Sheets("Discrete").Select
End Sub

Private Sub Wujiang()
    Sheets("Wujiang").Select
End Sub
