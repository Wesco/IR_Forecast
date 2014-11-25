Attribute VB_Name = "Imports"
Option Explicit

Sub ImportForecast()
    ImportFile "Campbellsville", Sheets("Cville").Range("A1")
    ImportFile "DLC", Sheets("DLC").Range("A1")
    ImportFile "Unicov", Sheets("Unicov").Range("A1")
    ImportFile "Mox BB", Sheets("Mox BB").Range("A1")
    ImportFile "Discrete", Sheets("Discrete").Range("A1")
    ImportFile "Wujiang", Sheets("Wujiang").Range("A1")
End Sub

Private Sub ImportFile(Forecast As String, Destination As Range)
    If MsgBox("Import " & Forecast & "?", vbYesNo, "Import File") = vbYes Then
        UserImportFile Destination, DelFile:=False, ShowAllData:=True, FileFilter:=Forecast & " (*.*), *.*", Title:="Import " & Forecast
    End If
End Sub
