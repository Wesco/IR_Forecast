Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatForecast()
    Format_Cville
End Sub

Private Sub Format_Cville()
    Dim TotalCols As Integer
    Dim TotalRows As Long

    Sheets("Cville").Select
    TotalCols = Rows(2).Columns(Columns.Count).End(xlToLeft).Column

    'Remove report description
    Rows(1).Delete
    
    'Remove unvalidated columns
    Columns("H:" & Split(Columns(TotalCols).Address(False, False), ":")(0)).Delete
    
    'Remove Description and Supplier Name
    Columns("B:C").Delete
    
    'Fill blanks with 0's
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("B2:E" & TotalRows).SpecialCells(xlCellTypeBlanks).Value = 0
End Sub
