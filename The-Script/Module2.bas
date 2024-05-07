'Run subtotals across worksheets 


Sub RunAllSubtotals()
    Dim sheetNames As Variant
    Dim sheetName As Variant
    
    Application.DisplayAlerts = False
    
    sheetNames = Array("", "", "", "", "") 'List of sheet names to process - amend as needed
    
    For Each sheetName In sheetNames
        Select Case sheetName
            Case "" 'Add relevant worksheet name
                Sub1 sheetName
            Case ""
                Sub2 sheetName
            Case ""
                Sub3 sheetName
            Case ""
                Sub4 sheetName
            Case ""
                Sub5 sheetName
         'Add more cases if needed
        End Select
    Next sheetName
    
    Application.DisplayAlerts = True
    
End Sub

Sub Sub1(sheetName As Variant)

   Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    
    Set ws = ActiveWorkbook.Worksheets(sheetName) 'Set worksheet

    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row 'Identify last row (amend as needed) before sorting
    
    ws.Range("A:P").RemoveSubtotal
    
    'Sort by Column E (this column will change) in ascending order
    'Change Header = xlNo if needed
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("E1:E" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:P" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row 'Identify the new last row after sorting

    'Add subtotals with each change in Column E for the data in Columns A to P (change as needed)
    'GroupBy:= will need to be changed as per your needs
    ws.Range("I6").Subtotal GroupBy:=5, Function:=xlSum, TotalList:=Array(3), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    Set dataRange = ws.Range("A2:P" & lastRow)

End Sub

Sub Sub2(sheetName As Variant)

        Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range

    Set ws = ActiveWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ws.Range("A:P").RemoveSubtotal
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("I1:I" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:P" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ws.Range("I6").Subtotal GroupBy:=9, Function:=xlSum, TotalList:=Array(3), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    Set dataRange = ws.Range("A2:P" & lastRow)

End Sub

Sub Sub3(sheetName As Variant)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range

    Set ws = ActiveWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ws.Range("A:P").RemoveSubtotal

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("D1:D" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:P" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ws.Range("I6").Subtotal GroupBy:=4, Function:=xlSum, TotalList:=Array(3), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    Set dataRange = ws.Range("A2:P" & lastRow)
End Sub

Sub Sub4(sheetName As Variant)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range

    Set ws = ActiveWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ws.Range("A:P").RemoveSubtotal

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("B1:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:P" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ws.Range("I6").Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(3), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    Set dataRange = ws.Range("A2:P" & lastRow)
End Sub

Sub Sub5(sheetName As Variant)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range

    Set ws = ActiveWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ws.Range("A:P").RemoveSubtotal

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("G1:G" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:P" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ws.Range("I6").Subtotal GroupBy:=7, Function:=xlSum, TotalList:=Array(3), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    Set dataRange = ws.Range("A2:P" & lastRow)

End Sub

