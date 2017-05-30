Sub caUploadFileInitialize(adjustmentBook)
    
    adjustmentBook.Sheets("sheet1").Range("a1") = "Auction Title"
    adjustmentBook.Sheets("sheet1").Range("b1") = "Inventory Number"
    adjustmentBook.Sheets("sheet1").Range("c1") = "Quantity Update Type"
    adjustmentBook.Sheets("sheet1").Range("d1") = "Quantity"
    adjustmentBook.Sheets("sheet1").Range("e1") = "Flag"
    adjustmentBook.Sheets("sheet1").Range("f1") = "FlagDescription"

End Sub

Sub caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)

    Dim lastRow As Long
    lastRow = adjustmentBook.Sheets("sheet1").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).row
    
    adjustmentBook.Sheets("sheet1").Range("B" & lastRow).Value = sku
    adjustmentBook.Sheets("sheet1").Range("C" & lastRow).Value = adjustmentType
    adjustmentBook.Sheets("sheet1").Range("D" & lastRow).Value = adjustment
    adjustmentBook.Sheets("sheet1").Range("E" & lastRow).Value = flag
    adjustmentBook.Sheets("sheet1").Range("F" & lastRow).Value = flagDescription

End Sub



