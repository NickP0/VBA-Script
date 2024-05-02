'Import data
Public Sub Import()
    Dim FileImport As Workbook
    Dim OpenFiles As Variant
    Dim targetWorksheets As Variant
    Dim i As Integer
    Dim targetSheet As Worksheet
    
    targetWorksheets = Array("", "", "", "", "", "raw") 'Amend as needed

    
    ' Use GetOpenFilename method to prompt for file selection
    OpenFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=False) 'Change MultiSelect:=True if needed
    ' MultiSelect set to false as should be importing 1 Ramp csv file

    Application.ScreenUpdating = False
    
    If VarType(OpenFiles) = vbBoolean Then
        MsgBox ""
        Exit Sub
    End If
    
    If OpenFiles <> "False" Then
        Set FileImport = Workbooks.Open(OpenFiles)
        
        For i = LBound(targetWorksheets) To UBound(targetWorksheets)
            Set targetSheet = ThisWorkbook.Worksheets(targetWorksheets(i))
            targetSheet.Activate     
            targetSheet.Range("A:P").RemoveSubtotal 'Change range as needed
            ThisWorkbook.Worksheets(targetWorksheets(i)).Range("A1").Select
            FileImport.Sheets(1).Range("A1").CurrentRegion.Copy  
            ActiveSheet.Paste
            Application.CutCopyMode = False
        Next i
        
        FileImport.Close
    Else
        MsgBox "No valid files selected."
    End If
    
    Application.ScreenUpdating = True
End Sub




