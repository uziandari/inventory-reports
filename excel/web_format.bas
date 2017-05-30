Sub SaveFileToImport(fileName As String, Optional filetype As Integer = 6, Optional filePath As String = "")

    If filePath = "" Then
        filePath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\database_files\inventory\import_files\"
    End If

    Dim NewName As String

    NewName = fileName

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=filePath & NewName, FileFormat:=filetype, CreateBackup:=False
    Application.DisplayAlerts = True

End Sub
Sub FormatReport()

  If Left(Trim(ActiveSheet.Name), 15) = "skus-" & Format(Now, "yyyy-MM-dd") Then
    Call CopytoJD
  End If

  Select Case Left(Trim(ActiveSheet.Name), 13)
    Case "ICAdjustments"
      Call AdjFormat
    Case "ICPreviousLoc"
      Call PreviousLocationFormat
    Case "ICPreviousPro"
      Call PreviousUpcFormat
    Case "ICAllReceipts"
      Call PreviousReceiptsFormat
    Case "ICDailySalesR"
      Call SalesFormat
    Case "InventoryExpo"
      If Cells(Rows.Count, "A").End(xlUp).Row > 5000 Then
        Call CAFormat
      Else
        Call SaveFileToImport("daily_b2b")
      End If
    Case "Inventory"
      Call NSFormat
    Case "ICReceiptDate"
      Call ReceiptFormat
    Case "ICCostResults"
      Call CostFormat
    Case "report34"
      Call CopytoBucket
    Case "report828"
      Call CopytoWholesalePending
    Case "ICDropshipRes"
      Call CopytoDropship
    Case "ns adj and In"
      Call CopytoDelists
      Call CopytoRelists
  End Select

End Sub
Sub AdjFormat()

    If Left(Trim(ActiveSheet.Name), 20) = "ICAdjustmentsResults" Then
        ActiveSheet.Name = "ICAdjustments"
        
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row

        Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
        
        Range("C2:C" & lastRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = 0
        
        
        Application.DisplayAlerts = False
        Call SaveFileToImport("adj_import")
        Application.DisplayAlerts = True
    Else
        Exit Sub
    End If
    
End Sub
Sub PreviousLocationFormat()

    If Left(Trim(ActiveSheet.Name), 19) = "ICPreviousLocations" Then
        ActiveSheet.Name = "ICPreviousLocations"
        
        Cells(1, 1).Value = "sku"
        Cells(1, 2).Value = "date"
        Cells(1, 3).Value = "field"
        Cells(1, 4).Value = "location"
        
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row

        Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
        
        
        Application.DisplayAlerts = False
        Call SaveFileToImport("prev_location_import")
        Application.DisplayAlerts = True
    Else
        Exit Sub
    End If
    
End Sub
Sub PreviousUpcFormat()

    If Left(Trim(ActiveSheet.Name), 22) = "ICPreviousProductCodes" Then
        ActiveSheet.Name = "ICPreviousProductCodes"
        
        Cells(1, 1).Value = "sku"
        Cells(1, 2).Value = "date"
        Cells(1, 3).Value = "upc"
        
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        Range("C2:C" & lastRow).NumberFormat = "0"
        Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
        
        Application.DisplayAlerts = False
        Call SaveFileToImport("prev_upc_import")
        Application.DisplayAlerts = True
    Else
        Exit Sub
    End If
    
End Sub
Sub PreviousReceiptsFormat()

    If Left(Trim(ActiveSheet.Name), 20) = "ICAllReceiptsResults" Then
        ActiveSheet.Name = "ICPreviousReceipts"
        
        Columns("A:A").Delete Shift:=xlToLeft
        
        Cells(1, 1).Value = "document"
        Cells(1, 2).Value = "date"
        Cells(1, 3).Value = "sku"
        Cells(1, 4).Value = "quantity"
        Cells(1, 5).Value = "type"
        
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
        
        Range("D2:D" & lastRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "NULL"
        
        Application.DisplayAlerts = False
        Call SaveFileToImport("prev_receipt_import")
        Application.DisplayAlerts = True
    Else
        Exit Sub
    End If
    
End Sub
Sub SalesFormat()

    If Left(Trim(ActiveSheet.Name), 19) = "ICDailySalesResults" Then
        ActiveSheet.Name = "ICDailySales"
        
        Dim ws As Worksheet
        Set ws = Worksheets("ICDailySales")
        
        Dim lastRow As Long
        Dim rng As Range
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row

        Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
        
        Set rng = ws.Range("C1:C" & lastRow)
        With rng
            .AutoFilter Field:=1, Criteria1:="<>*@*"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With
    
        ws.AutoFilterMode = False
        
        Columns("A:A").Delete Shift:=xlToLeft
        
        Application.DisplayAlerts = False
        Call SaveFileToImport("daily_sales_import")
        Application.DisplayAlerts = True
    Else
        Exit Sub
    End If
    
End Sub


Sub CAFormat()

    If ActiveSheet.Name = "InventoryExport" Then
    
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        ActiveSheet.Range("a1").CurrentRegion.Select
            
            Columns("C:D").Delete Shift:=xlToLeft
            Columns("B:B").Cut
            Columns("A:A").Insert Shift:=xlToRight
            Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1))
            Range("C1").FormulaR1C1 = "Available"
            Columns("C:C").Cut
            Columns("E:E").Insert Shift:=xlToRight
            Columns("D:D").SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "NULL"
            
            Application.DisplayAlerts = False
            
            Columns("J").Cut
            Columns("M").Insert Shift:=xlToRight
            
            Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1), Array(3, 1), Array(4, 1)), _
                TrailingMinusNumbers:=True
            Columns("M:AA").Delete Shift:=xlToLeft
            Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
                "=", FieldInfo:=Array(Array(1, 1), Array(2, 9)), TrailingMinusNumbers:=True
            Columns("M:U").Delete Shift:=xlToLeft
            Range("L1").FormulaR1C1 = "img"
            
            Application.DisplayAlerts = True
     
        
            Call SaveFileToImport("ca_import")
        
    
    End If
    
End Sub

Sub NSFormat()

    If ActiveSheet.Name = "Inventory" Then
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("E2:E" & lastRow).NumberFormat = "0"
    
    Columns("K:AA").Delete Shift:=xlToLeft
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1))
    Range("F1").FormulaR1C1 = "On Hand"

    Columns("G:G").TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1))
    Range("G1").FormulaR1C1 = "Committed"
    
    Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1))
    Range("J1").FormulaR1C1 = "Purchase Price"
    Range("J2:J" & lastRow).NumberFormat = "0.00"
    Range("J2:J" & lastRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "NULL"

    Cells(1, 1).Select
    
    Dim filePath As String
    Dim NewName As String

    Call SaveFileToImport("ns_import")

    End If

End Sub

Sub ReceiptFormat()

    If Left(Trim(ActiveSheet.Name), 13) = "ICReceiptDate" Then
        ActiveSheet.Name = "LastItemReceipt"
    Else
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets("LastItemReceipt")
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
    
    
    Call SaveFileToImport("receipt_date_import")

End Sub

Sub CostFormat()

    If Left(Trim(ActiveSheet.Name), 13) = "ICCostResults" Then
        ActiveSheet.Name = "ItemCost"
    Else
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets("ItemCost")
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    Range("B2:C" & lastRow).NumberFormat = "0.00"
    Range("B2:C" & lastRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "NULL"
    
    
    Call SaveFileToImport("item_cost_import")

End Sub

Sub CopytoRelists()

    Dim Msg As String, Ans As Variant
     
     
    Msg = "Include Returns?"
      
    Ans = MsgBox(Msg, vbYesNo)
    
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    
    Set ws = Workbooks("NS ADJ.xlsx").Worksheets("ns adj and Inline Delists")
    Set ws2 = Workbooks("NS ADJ.xlsx").Worksheets("relist")
      
    Select Case Ans
          
    Case vbYes
        
              
        lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        lastRow2 = ws2.Range("A" & ws.Rows.Count).End(xlUp).Row

        Set rng = ws.Range("A1:C" & lastRow)
          
        Set NewBook = Workbooks.Add
        
        With rng
            .AutoFilter Field:=3, Criteria1:=Array("=", Date)
            .SpecialCells(xlCellTypeVisible).Copy
            NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
        End With
        
        Dim relistBookLastRow As Long
        relistBookLastRow = NewBook.Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
        Workbooks("NS ADJ.xlsx").Worksheets("relist").Range("A2:A" & lastRow2).Copy
        NewBook.Worksheets("Sheet1").Range("A" & relistBookLastRow).PasteSpecial (xlPasteValues)
        NewBook.Worksheets("Sheet1").Columns("B:I").EntireColumn.Delete
        ws.AutoFilterMode = False
        
          
    Case vbNo

        lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row

        Set rng = ws.Range("A1:C" & lastRow)
          
        Set NewBook = Workbooks.Add
        
        With rng
            .AutoFilter Field:=3, Criteria1:=Array("=", Date)
            .SpecialCells(xlCellTypeVisible).Copy
            NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
            NewBook.Worksheets("Sheet1").Columns("B:C").EntireColumn.Delete
        End With
    
        ws.AutoFilterMode = False

    End Select

    Call SaveFileToImport("relist")

End Sub

Sub CopytoDelists()
    
    Dim lastRow As Long
    lastRow = Workbooks("NS ADJ.xlsx").Worksheets("ns adj and Inline Delists").Cells(Rows.Count, "A").End(xlUp).Row
    
    
    Set NewBook = Workbooks.Add
    Workbooks("NS ADJ.xlsx").Worksheets("ns adj and Inline Delists").Range("A1:A" & lastRow).Copy
    NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
    
    If Not IsEmpty(Workbooks("NS ADJ.xlsx").Worksheets("relist").Range("A2").Value) Then
        Dim relistBookLastRow As Long
        relistBookLastRow = NewBook.Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
        lastRow = Workbooks("NS ADJ.xlsx").Worksheets("relist").Cells(Rows.Count, "A").End(xlUp).Row
        
        Workbooks("NS ADJ.xlsx").Worksheets("relist").Range("A2:A" & lastRow).Copy
        NewBook.Worksheets("Sheet1").Range("A" & relistBookLastRow).PasteSpecial (xlPasteValues)
    
    End If
    
    If Not IsEmpty(Workbooks("NS ADJ.xlsx").Worksheets("Relist NA's").Range("A2").Value) Then
        Dim naBookLastRow As Long
        naBookLastRow = NewBook.Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
        lastRow = Workbooks("NS ADJ.xlsx").Worksheets("Relist NA's").Cells(Rows.Count, "A").End(xlUp).Row
    
        Workbooks("NS ADJ.xlsx").Worksheets("Relist NA's").Range("A2:A" & lastRow).Copy
        NewBook.Worksheets("Sheet1").Range("A" & naBookLastRow).PasteSpecial (xlPasteValues)
    
    End If
    
    Call SaveFileToImport("prevdelists")

End Sub


Sub CopytoBucket()
    
    Dim lastRow As Long
    lastRow = Workbooks("Bucket").Worksheets("report34").Cells(Rows.Count, "A").End(xlUp).Row
    
    
    Set NewBook = Workbooks.Add
    Workbooks("Bucket").Worksheets("report34").Range("E1:E" & lastRow).Copy
    NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
    
    
    Call SaveFileToImport("daily_bucket")

End Sub
Sub CopytoWholesalePending()

    Dim lastRow As Long
    lastRow = Workbooks("Wholesale").Worksheets("report828").Cells(Rows.Count, "A").End(xlUp).Row
    
    
    Set NewBook = Workbooks.Add
    Workbooks("Wholesale").Worksheets("report828").Range("A1:A" & lastRow).Copy
    NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
    
    Dim newLastRow As Long
    newLastRow = NewBook.Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row

    ActiveSheet.Range("A1:A" & newLastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    Call SaveFileToImport("wholesale_committed")

End Sub


Sub CopytoJD()

    If Left(Trim(ActiveSheet.Name), 15) = "skus-" & Format(Now, "yyyy-MM-dd") Then
        
        Columns("B:H").EntireColumn.Delete
        
        Call SaveFileToImport("daily_bc")

    Else
        Exit Sub
    End If

End Sub

Sub CopytoDropship()

    If Left(Trim(ActiveSheet.Name), 17) = "ICDropshipResults" Then
        ActiveSheet.Name = "ICDropshipResults"
    Else
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets("ICDropshipResults")
    
    Columns("A:A").Delete Shift:=xlToLeft
    Columns("B:B").Delete Shift:=xlToLeft

    Call SaveFileToImport("dropship_skus")

End Sub



