'This procedure will enable a pop up everytime the workbook is opened
Private Sub Workbook_Open()
    MsgBox prompt:="Welcome to " & ActiveWorkbook.Name, Buttons:=vbExclamation
    MsgBox prompt:="Please do not amend any values within this Dashboard", Buttons:=vbExclamation 'Change as needed
    MsgBox prompt:="If you encounter any issues, reach out to Nick", Buttons:=vbExclamation
End Sub

'This procedure will detect any changes made on the pivot tables datasource and amend the pivot table accordingly
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    Dim pt As PivotTable
    Dim ptrange As Range
    Dim irange As Range

' turn EnableEvents off to prevent infinite looping of the event change - will crash if this is set to True 
Application.EnableEvents = False

Application.ScreenUpdating = False

    On Error Resume Next
    ' Pivot Table Object
    Set pt = Worksheets("").PivotTables("")
    
    ' Pivot Table Data Source
    'Set ptrange = pt.SourceData:=""
    Set ptrange = Worksheets(# INSERT WORKSHEET HERE #).Range("")
    
    
    Set irange = Intersect(ptrange, target)
    
    If ptrange Is Nothing Then
        MsgBox "", vbExclamation
        Exit Sub
    End If
    
    If IsEmpty(irange) = False Then
        Application.ScreenUpdating = False
        pt.PivotCache.Refresh
        Application.ScreenUpdating = True
    End If

Application.ScreenUpdating = True

' turn EnableEvents back on
Application.EnableEvents = True
End Sub
