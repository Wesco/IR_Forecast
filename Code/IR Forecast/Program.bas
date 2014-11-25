Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    ImportForecast
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean
    Dim s As Worksheet
    
    PrevScrnUpdat = Application.ScreenUpdating
    PrevDispAlert = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Columns.Hidden = False
            s.Rows.Hidden = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next
    
    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat
End Sub
