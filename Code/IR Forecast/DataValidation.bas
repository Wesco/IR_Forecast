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
    Dim TotalCols As Integer
    Dim dt As Date
    Dim i As Integer

    Sheets("Campbellsville").Select
    If Not ActiveSheet.UsedRange.Rows.Count > 1 Then Exit Sub
    TotalCols = Rows(2).Columns(Columns.Count).End(xlToLeft).Column

    'Check the first 3 column headers
    For i = 0 To 2
        If Cells(2, i + 1).Value <> Array("Part #", "Part Description", "Supplier Name")(i) Then
            Err.Raise CustErr.COLNOTFOUND, "Campbellsville", "Report validation failure."
        End If
    Next

    'Check columns D:G and make sure they are dates
    For i = 4 To 7
        'This will throw an error if the column found is not a date
        TypeName CDate(Cells(2, i).Value)
    Next
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
