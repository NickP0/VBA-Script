'Random - deprecated
'ThisWorkbook.cls will be better to use
'Can implement this into a button if needed

Sub RefreshPivotDataSource()

    Range("B7").Select
    ActiveSheet.PivotTables("PivotTable5").PivotCache.Refresh
End Sub
