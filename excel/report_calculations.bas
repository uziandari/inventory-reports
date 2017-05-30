Sub addToSpecial(sku)

    Dim specialRow As Long
    speicalRow = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("Special").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    Sheets("Special").Range("A" & speicalRow).Value = sku

End Sub

Sub CalculateLessThanNine()
    
    Set lessNineSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("less_nine")
  
    Dim lastRow As Long
    lastRow = lessNineSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)
    
    'return to proper worksheet
    lessNineSheet.Activate
    
    ' loop through input
    For i = 2 To lastRow
        Dim sku As String
        
        If Len(Trim(Cells(i, 11).Value)) = 0 Then
            Cells(i, 12).Value = "Needs Input"
        ElseIf Not IsNumeric(Cells(i, 11)) Then
            Cells(i, 12).Value = "Look Here"
        Else
            Dim stock As Integer, committed As Integer, pending As Integer, available As Integer, actual As Integer
        
            actual = Cells(i, 11).Value
            stock = Cells(i, 10).Value
            committed = Cells(i, 6).Value
            pending = Cells(i, 5).Value
            available = Cells(i, 4).Value

            ' call function to calculate adjustment needed
            adjustment = calculateLessNineAdjustment(actual, stock, committed, pending, available)

            Cells(i, 12).Value = adjustment

            Dim adjustmentType As Variant, flag As String, flagDescription As String

            If IsNumeric(adjustment) Then
                sku = Cells(i, 1).Value
                adjustmentType = "Relative"
                flag = "GreenFlag"
                flagDescription = "final qty " & Format(CStr(Now), "m/d")

                Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)
            ElseIf adjustment = "Make Zero" Then
                sku = Cells(i, 1).Value
                adjustmentType = "Absolute"
                flag = "GreenFlag"
                flagDescription = "final qty " & Format(CStr(Now), "m/d")
                adjustment = 0

                Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)

            ElseIf adjustment = "ok" Then
                sku = Cells(i, 1).Value
                adjustmentType = Null
                adjustment = Null
                flag = "GreenFlag"
                flagDescription = "final qty " & Format(CStr(Now), "m/d")

                Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)
            End If
            
        End If
    Next i
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("lessnineCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")

    
End Sub


Function calculateLessNineAdjustment(actual, stock, committed, pending, available) As Variant
    
    
    If actual > stock Then
        actual = stock
    End If
    
    If actual = 1 And committed + pending = 0 Then
        calculateLessNineAdjustment = acutal - available + 1
        Exit Function
    End If
    
    If actual - (committed + pending) <= 0 Then
        calculateLessNineAdjustment = "Make Zero"
        Exit Function
    ElseIf actual - (committed + pending) > 5 Then
        actual = actual - committed - pending - 2
    Else
        actual = actual - committed - pending - 1
    End If

    If actual - available >= 0 Then
        calculateLessNineAdjustment = "ok"
    Else
        calculateLessNineAdjustment = actual - available
    End If
End Function

Function calculateRelistAdjustment(actual, inline) As Integer

    If actual = 0 Then
        calculateRelistAdjustment = 0
    ElseIf actual = 1 Then
        calculateRelistAdjustment = 1
    ElseIf actual <= 3 And inline = "Yes" Then
        calculateRelistAdjustment = actual
    ElseIf actual <= 5 And actual > 1 Then
        calculateRelistAdjustment = actual - 1
    Else
        calculateRelistAdjustment = actual - 2
    End If
    
End Function

Function calculateFlag(available, inline, sku) As String
    
    If inline = "Yes" Then
        calculateFlag = "BlueFlag"
    ElseIf available > 0 And available < 13 Then
        calculateFlag = "GreenFlag"
    ElseIf available > 12 Then
        calculateFlag = "NoFlag"
    Else '0
        calculateFlag = "RedFlag"
    End If

End Function

Function calculateFlagDescription(flag, reportString, available) As String

    If flag = "BlueFlag" Then
        calculateFlagDescription = "Inline"
    ElseIf flag = "GreenFlag" Then
        calculateFlagDescription = "final qty " & Format(CStr(Now), "m/d") & " " & reportString
    'ElseIf flag = "RedFlag" And available > 0 Then
    '    calculateFlagDescription = "absolute ebay final " & Format(CStr(Now), "m/d")
    ElseIf flag = "RedFlag" And available = 0 Then
        calculateFlagDescription = "absolute final " & Format(CStr(Now), "m/d/yy")
    Else 'NoFlag
        calculateFlagDescription = "_DELETE_"
    End If

End Function

Sub CalculateRelist()

    Set relistSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("relist")
  
    Dim lastRow As Long
    lastRow = relistSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)
    
    'return to proper worksheet
    relistSheet.Activate

    Cells(1, 13).Value = "CA Actual"

    ' loop through input
    For i = 2 To lastRow
        If Len(Trim(Cells(i, 8).Value)) = 0 Then
            Cells(i, 13).Value = "Needs Input"
            MsgBox("You need to provide data for row:" & i)
            Exit Sub
        ElseIf Not IsNumeric(Cells(i, 8)) Then
            Cells(i, 13).Value = "Look Here"
        Else
            Dim actual As Integer
        
            actual = Cells(i, 8).Value
            inline = Cells(i, 10).Value
            
            ' call function to calculate adjustment needed
            adjustment = calculateRelistAdjustment(actual, inline)
            
            Cells(i, 13).Value = adjustment
            
            Dim sku As String
            
            sku = Cells(i, 1).Value
            
            'Finds updated locations --needs attention
'            If Not IsEmpty(Cells(i, 9)) Then
'                Dim newLocation As String
'                newLocation = Cells(i, 9)
'                Call upload_files.LocationUpdate(sku, newLocation)
'            End If
            'Ends location update
            
            
            adjustmentType = "Absolute"
            flag = calculateFlag(adjustment, inline, sku)
            flagDescription = calculateFlagDescription(flag, "(wr)", adjustment)
            
            Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)
            
        End If
    Next i
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("relistCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")

End Sub

Function calculateAlertsAdjustment(actual, stock, committed, pending, available) As Variant
    
    
    If actual > stock Then
        actual = stock
    End If
    
    If actual = 1 And committed + pending = 0 Then
        calculateAlertsAdjustment = 1
        Exit Function
    End If
    
    If actual <= 0 Or actual - committed <= 0 Then
        calculateAlertsAdjustment = "delist"
    ElseIf actual - (committed + pending) <= 0 Then
        calculateAlertsAdjustment = 0
    ElseIf actual - (committed + pending) > 5 Then
        calculateAlertsAdjustment = actual - committed - pending - 2
    Else
        calculateAlertsAdjustment = actual - committed - pending - 1
    End If

    
End Function
Function calculateAlertsFlag(adjustment, inline) As String
    
    If inline = "Yes" Then
        calculateAlertsFlag = "BlueFlag"
    ElseIf adjustment < 13 Then
        calculateAlertsFlag = "GreenFlag"
    Else 'available/adjustment > 12
        calculateAlertsFlag = "NoFlag"
    End If

End Function
Function calculateAlertsFlagDescription(flag, reportString) As String

    If flag = "BlueFlag" Then
        calculateAlertsFlagDescription = "Inline"
    ElseIf flag = "GreenFlag" Then
        calculateAlertsFlagDescription = "final qty " & Format(CStr(Now), "m/d") & " " & reportString
    Else 'NoFlag
        calculateAlertsFlagDescription = "_DELETE_"
    End If

End Function
Sub CalculateAlerts()

    Set alertsSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("alerts")
  
    Dim lastRow As Long
    lastRow = alertsSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)
    
    'return to proper worksheet
    alertsSheet.Activate
    
    ' loop through input
    For i = 2 To lastRow
    
        Dim stock As Integer, committed As Integer, pending As Integer, available As Integer, actual As Variant
        Dim sku As String, description As String, location As String, backstock As String, inline As String, initialFlag As String
        
        actual = Cells(i, 11).Value
        stock = Cells(i, 10).Value
        committed = Cells(i, 6).Value
        pending = Cells(i, 5).Value
        available = Cells(i, 4).Value
        
        'add 0 to delist
        sku = Cells(i, 1).Value
        description = Cells(i, 2).Value
        inline = Cells(i, 13).Value
        
        If Cells(i, 12).Interior.ColorIndex <> 0 Then
            Cells(i, 12).Interior.ColorIndex = 0
        End If
    
        If Len(Trim(Cells(i, 11).Value)) = 0 Then
            Cells(i, 12).Value = "Needs Input"
            'Highlights adjustment cell red and exits the sub -- All input cells need values for routine to complete
            Cells(i, 12).Interior.ColorIndex = 3
            Exit Sub
        ElseIf Not IsNumeric(Cells(i, 11)) Then
            Dim skuSpecial As String
            skuSpecial = Cells(i, 1).Value
            Call addToSpecial(skuSpecial)
        Else

            ' call function to calculate adjustment needed
            adjustment = calculateAlertsAdjustment(actual, stock, committed, pending, available)

            Cells(i, 12).Value = adjustment
            
            If Cells(i, 12).Value = "delist" Then
                Call AddToDelist(sku, description, stock, committed, inline)
            End If

            
            
            'add to CA upload file
            If IsNumeric(Cells(i, 12)) Then
                Dim flag As String, flagDescription As String
                'call flag function to determine flag --only for items added to CA upload (no delists)
                flag = calculateAlertsFlag(adjustment, inline)
                flagDescription = calculateAlertsFlagDescription(flag, "(a)")
                Call upload_files.caUploadFileAdd(adjustmentBook, sku, "Absolute", adjustment, flag, flagDescription)
            End If
        End If

    Next i
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("alertsCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")

End Sub

Sub AddToDelist(sku, description, stock, committed, inline)

    Dim delistRow As Long
    delistRow = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("delist").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    Sheets("delist").Range("A" & delistRow).Value = sku
    Sheets("delist").Range("B" & delistRow).Value = description
    Sheets("delist").Range("C" & delistRow).Value = stock - committed
    Sheets("delist").Range("D" & delistRow).Value = inline
End Sub


Sub CalculateDelist()

    Set delistSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("delist")

    Dim lastRow As Long
    lastRow = delistSheet.Cells(Rows.Count, "A").End(xlUp).row

    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)

    'return to proper worksheet
    delistSheet.Activate


    'Gathers recount date
    Dim recountDay As Date, recountString As String

    recountDay = DateAdd("d", 2, Now())
    
    If Weekday(recountDay) = 7 Then
        recountDay = DateAdd("d", 4, Now())
    ElseIf Weekday(recountDay) = 1 Then
        recountDay = DateAdd("d", 3, Now())
    
    End If
    
    recountString = Format(CStr(recountDay), "m/d")
    
    'End date gathering
    
    For i = 2 To lastRow
    
        Dim sku As String, inline As String, flag As String, flagDescription As String
        Dim adjustmentType As Variant, adjustment As Variant
        
        'add 0 to delist
        sku = Cells(i, 1).Value
        inline = Cells(i, 4).Value
 
        adjustmentType = Null
        adjustment = Null
        If inline = "Yes" Then
            flag = "BlueFlag"
            flagDescription = "Inline"
        Else 'non-inline SKUs
            flag = "YellowFlag"
            flagDescription = "final recount " & recountString
        End If
        
        Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)
    
    Next i
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("delistCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")


End Sub
Sub inlineDatePush(inlineAdjustSheet, i)
    
    'Gathers day of the week
    Dim recountDay As Date
    
    inlineAdjustSheet.Activate
    
    recountDay = inlineAdjustSheet.Cells(i, 3).Value

    inlineAdjustSheet.Cells(i, 3).Value = Format(DateAdd("d", 1, recountDay), "MM/dd/yyyy")

    If Weekday(inlineAdjustSheet.Cells(i, 3).Value) = 3 Then
        inlineAdjustSheet.Cells(i, 3).Value = Format(DateAdd("d", 2, recountDay), "MM/dd/yyyy")
    ElseIf Weekday(inlineAdjustSheet.Cells(i, 3).Value) = 7 Then
        inlineAdjustSheet.Cells(i, 3).Value = Format(DateAdd("d", 3, recountDay), "MM/dd/yyyy")
    End If

End Sub


Sub calculateRelistToPush()

    Application.ScreenUpdating = False

    Set pushedSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("relist_pushed")
    Set inlineAdjustSheet = Workbooks("NS ADJ").Sheets("ns adj and Inline Delists")
    
    Dim lastRow As Long
    Dim inlineLastRow As Long
    
    lastRow = pushedSheet.Cells(Rows.Count, "A").End(xlUp).row
    inlineLastRow = inlineAdjustSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    
    For i = 2 To inlineLastRow
        Dim sku As String
        
        sku = inlineAdjustSheet.Cells(i, 1).Value
        
        If pushedSheet.Cells(2, 1) = "" Then
            Exit Sub
        End If
        
        'work here for proper loop
        Set rgFound = pushedSheet.Range("A1:A" & lastRow).Find(sku)
        
        If Not rgFound Is Nothing Then
           Call inlineDatePush(inlineAdjustSheet, i)
        End If
        
        
    Next i
    
    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)
    
    pushedSheet.Activate

    For k = 2 To lastRow
    
        If Cells(k, 4).Value > 0 Then
            Cells(k, 1).Interior.ColorIndex = 8
        End If
        
        If Cells(k, 13).Value <> "Yes" Then
            Dim pushedSku As String, flag As String, flagDescription As String, recountString As String
            Dim adjustmentType As Variant, adjustment As Variant
        
            pushedSku = Cells(k, 1).Value
            
            recountString = Format(DateAdd("d", 1, Now()), "m/d")

            If Weekday(recountString) = 3 Then
                recountString = Format(DateAdd("d", 2, Now()), "m/d")
            ElseIf Weekday(recountString) = 7 Then
                recountString = Format(DateAdd("d", 3, Now()), "m/d")
            End If
 
            adjustmentType = Null
            adjustment = Null
            
            flag = "YellowFlag"
            flagDescription = "final recount " & recountString & " (pushed)"
        
            Call upload_files.caUploadFileAdd(adjustmentBook, pushedSku, adjustmentType, adjustment, flag, flagDescription)
    
        End If
        
    Next k
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("pushedCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")
    
    Workbooks("pushedCA").Close

    Application.ScreenUpdating = True

End Sub

Sub calculateB2bInventory()

    Application.ScreenUpdating = False

    Set b2bSheet = Workbooks("Reports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("b2b_quantity")
    
    Dim lastRow As Long
    
    lastRow = b2bSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)
    
    b2bSheet.Activate
    
    b2bSheet.Range("A2:A" & lastRow).Copy
    adjustmentBook.Sheets("sheet1").Range("B2").PasteSpecial Paste:=xlPasteValues
    adjustmentBook.Sheets("sheet1").Range("C2:C" & lastRow).Value = "absolute"
    b2bSheet.Range("B2:B" & lastRow).Copy
    adjustmentBook.Sheets("sheet1").Range("D2").PasteSpecial Paste:=xlPasteValues
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("b2bUpload", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")
    
    Workbooks("b2bUpload").Close

    Application.ScreenUpdating = True

End Sub
