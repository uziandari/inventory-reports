Sub GenericCAInventoryFormat()

    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    
    ActiveSheet.Range("a1").CurrentRegion.Select
        
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="=", FieldInfo:=Array(Array(1, 9), Array(2, 1))
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Available"
    Columns("C:C").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    
    Call db_format.FormatTable

End Sub
