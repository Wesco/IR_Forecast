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
    Dim TotalCols As Integer
    Dim dt As Date
    Dim i As Integer

    Sheets("DLC").Select
    If Not ActiveSheet.UsedRange.Rows.Count > 1 Then Exit Sub
    TotalCols = Rows(3).Columns(Columns.Count).End(xlToLeft).Column

    'Check the first 4 column headers
    For i = 0 To 3
        If Cells(3, i + 1).Value <> Array("Supplier Site", "Item", "Description", "Primary UOM")(i) Then
            Err.Raise CustErr.COLNOTFOUND, "DLC", "Report validation failure."
        End If
    Next

    'Check the remaining columns and make sure they are dates
    For i = 5 To TotalCols
        'This will throw an error if the column found is not a date
        TypeName CDate(Cells(2, i).Value)
    Next
End Sub

Private Sub Unicov()
    Dim TotalCols As Integer
    Dim dt As Date
    Dim i As Integer

    Sheets("Unicov").Select
    If Not ActiveSheet.UsedRange.Rows.Count > 1 Then Exit Sub
    TotalCols = Rows(3).Columns(Columns.Count).End(xlToLeft).Column

    'Check the first 5 column headers
    For i = 0 To 4
        If Cells(6, i + 1).Value <> Array("ITEM", "DESCRIPTION", "UOM", "SUPPLIER_NAME", "SUPPLIER_SITE_NAME")(i) Then
            Err.Raise CustErr.COLNOTFOUND, "Unicov", "Report validation failure."
        End If
    Next

    'Check the remaining columns and make sure they are dates
    For i = 5 To TotalCols
        'This will throw an error if the column found is not a date
        TypeName CDate(Cells(2, i).Value)
    Next
End Sub

Private Sub MoxBB()
    Dim TotalCols As Integer
    Dim dt As Date
    Dim i As Integer

    Sheets("Mox BB").Select
    If Not ActiveSheet.UsedRange.Rows.Count > 1 Then Exit Sub
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    
    'Check first 2 column headers
    For i = 0 To 1
        If Cells(1, i + 1).Value <> Array("Item", "Description")(i) Then
            Err.Raise CustErr.COLNOTFOUND, "Unicov", "Report validation failure."
        End If
    Next
    
    'Check the remaining columns and makre sure they are dates
    For i = 3 To TotalCols
        TypeName CDate(Cells(1, i).Value & "-" & Year(Date))
    Next
End Sub

Private Sub Discrete()
    Dim TotalCols As Integer
    Dim dt As Date
    Dim i As Integer
    
    Sheets("Discrete").Select
    If Not ActiveSheet.UsedRange.Rows.Count > 1 Then Exit Sub
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    
    'Check first 2 column headers
    For i = 0 To 1
        If Cells(1, i + 1).Value <> Array("Item", "Description")(i) Then
            Err.Raise CustErr.COLNOTFOUND, "Unicov", "Report validation failure."
        End If
    Next
    
    'Check the remaining columns and makre sure they are dates
    For i = 3 To TotalCols
        TypeName CDate(Cells(1, i).Value & "-" & Year(Date))
    Next
End Sub

Private Sub Wujiang()
    Sheets("Wujiang").Select
End Sub
