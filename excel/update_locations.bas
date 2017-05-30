Sub locationFileInitialize(locationBook)
    
    'fills header row
    locationBook.Sheets("Location Upload File").Range("A1:AA1") = Array("Item Name/Number", "Display Name/Code", "Parent", "Item Type", "Item Sub-Type", "Sales Description", "Purchase Description", "Price Level:<pricelevel>", "Income Account", "Asset Account", "COGS/Expense Account", "Costing Method", "Quantity on Hand", "Serial Numbers", "Reorder Point", "Preferred Vendor", "Purchase Price", "Drop Ship", "Tax Code", "Is Taxable", "Is Inactive", "Custom UPC", "Bin", "Backstock", "Headcover", "Inline", "Stock Description")


End Sub


Sub findNewLocations()

    Application.ScreenUpdating = False
    
    'assigns name to correct sheet in reports workbook
    On Error GoTo errorhandler:
    Set locationSheet = ActiveWorkbook.Sheets("moves")
  
    'finds last row from moves sheet
    Dim lastRow As Long
    lastRow = locationSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    'creates a new workbook and fills header row
    Set locationBook = Workbooks.Add
    locationBook.Sheets("sheet1").Name = "Location Upload File"
    
    Call locationFileInitialize(locationBook)
    
    'return to proper worksheet
    locationSheet.Activate
    
    For i = 2 To lastRow
    
        If Not Cells(i, 9).Value = vbNullString Then
            Dim sku As String, newLocation As String
            
            sku = Cells(i, 1).Value
            newLocation = Cells(i, 9).Value
            
            Call locationUploadFileAdd(locationBook, "Location Upload File", sku, newLocation)
            
        End If

    Next i
    
    'Save the new file
    Call saveLocationFile(locationBook)
    
    Application.ScreenUpdating = True
    
    Exit Sub
errorhandler:
    MsgBox ("Make sure you have the correct active workbook (the one with the 'moves' sheet).")
    
    

End Sub

Sub locationUploadFileAdd(adjustmentBook, locationSheet, sku, newLocation)

    Dim lastRow As Long
    lastRow = adjustmentBook.Sheets(locationSheet).Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    
    adjustmentBook.Sheets(locationSheet).Range("A" & lastRow).Value = sku
    adjustmentBook.Sheets(locationSheet).Range("D" & lastRow).Value = 1
    adjustmentBook.Sheets(locationSheet).Range("W" & lastRow).Value = newLocation

End Sub

Sub saveLocationFile(locationBook, Optional filetype As Integer = 6)
    
    locationBook.Sheets("Location Upload File").Activate
    
    Dim fileName As String, filePath As String

    fileName = "Location Upload File " & Format(CStr(Now), "yyyy_mm_dd")
    filePath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\"

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=filePath & fileName, FileFormat:=filetype, CreateBackup:=False
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Close

End Sub
