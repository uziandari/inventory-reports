Sub EmailReport(toList, ccList, subject, body, attachment)

    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = toList
        .CC = ccList
        .BCC = ""
        .subject = subject & Format(Date, "MM/DD/YYYY")
        .HTMLBody = body
        .Attachments.Add (attachment)
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing


End Sub

Sub EmailHighValueAdjustments()
    
    Dim today As String
    today = Format(Date, "mm/dd/yyyy")
    
    Dim compUserName As String
    compUserName = CreateObject("WScript.Network").Username

    If Format(FileDateTime("C:\Users\" & compUserName & "\Desktop\adjustments\ns_adj_hv.xlsx"), "mm/dd/yyyy") = today Then
        Call report_emails.EmailReport("email-list", "", "NS Adjustments (HV) ", "Attached.", "C:\Users\" & compUserName & "\Desktop\adjustments\ns_adj_hv.xlsx")
    End If
    
    
End Sub

Sub EmailDelist()
    
    Set delistSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("delist")

    Dim lastRow As Long, delistLastRow As Long, inlineLastRow As Long
    lastRow = delistSheet.Cells(Rows.Count, "A").End(xlUp).row

    Set DelistBook = Workbooks.Add
    Set InlineDelistBook = Workbooks.Add
    
    DelistBook.Sheets("Sheet1").Cells(1, 1).Value = "sku"
    DelistBook.Sheets("Sheet1").Cells(1, 2).Value = "description"
    DelistBook.Sheets("Sheet1").Cells(1, 3).Value = "ns qty"
    
    InlineDelistBook.Sheets("Sheet1").Cells(1, 1).Value = "sku"
    InlineDelistBook.Sheets("Sheet1").Cells(1, 2).Value = "description"
    InlineDelistBook.Sheets("Sheet1").Cells(1, 3).Value = "ns qty"
    
    delistLastRow = DelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    inlineLastRow = InlineDelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    
    'return to proper worksheet
    delistSheet.Activate
    
    Range("A2:D" & lastRow).Sort key1:=Range("A2:A" & lastRow), _
      order1:=xlAscending, Header:=xlYes
    
    For i = 2 To lastRow
        If Cells(i, 4).Value = "Yes" Or Cells(i, 4).Value = "Yes - No DS" Then
            Cells(i, 1).Copy Destination:=InlineDelistBook.Sheets("Sheet1").Range("A" & inlineLastRow)
            Cells(i, 2).Copy Destination:=InlineDelistBook.Sheets("Sheet1").Range("B" & inlineLastRow)
            Cells(i, 3).Copy Destination:=InlineDelistBook.Sheets("Sheet1").Range("C" & inlineLastRow)
            inlineLastRow = InlineDelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
        Else
            Cells(i, 1).Copy Destination:=DelistBook.Sheets("Sheet1").Range("A" & delistLastRow)
            Cells(i, 2).Copy Destination:=DelistBook.Sheets("Sheet1").Range("B" & delistLastRow)
            Cells(i, 3).Copy Destination:=DelistBook.Sheets("Sheet1").Range("C" & delistLastRow)
            delistLastRow = DelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
        End If
        
    Next i
    
    With DelistBook.Sheets("Sheet1").UsedRange
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 8
        .Interior.ColorIndex = xlNone
    End With
    
    With InlineDelistBook.Sheets("Sheet1").UsedRange
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 8
        .Interior.ColorIndex = xlNone
    End With
    
    DelistBook.Sheets("Sheet1").Columns("A:C").AutoFit
    InlineDelistBook.Sheets("Sheet1").Columns("A:C").AutoFit
    
    Dim delistEmail As Range
    Set delistEmail = DelistBook.Sheets("Sheet1").Range("A1:C" & delistLastRow)
    
    Dim InlineDelistEmail As Range
    Set InlineDelistEmail = InlineDelistBook.Sheets("Sheet1").Range("A1:C" & inlineLastRow)
    
    Call report_emails.EmailReport("email-list", "", "Inventory Delist " & Format(CStr(Now), "h:mm" & " "), RangetoHTML(delistEmail), "")
    Call report_emails.EmailReport("email-list", "", "Inline Inventory Delist " & Format(CStr(Now), "h:mm" & " "), RangetoHTML(InlineDelistEmail), "")
    
    DelistBook.Close SaveChanges:=False
    InlineDelistBook.Close SaveChanges:=False
    
End Sub

Sub RelistEmail()

    Set relistSheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("relist")

    Dim lastRow As Long, relistLastRow As Long, ebayRelistLastRow As Long
    lastRow = relistSheet.Cells(Rows.Count, "A").End(xlUp).row

    Set RelistBook = Workbooks.Add
    
    RelistBook.Sheets("Sheet1").Cells(1, 1).Value = "sku"
    RelistBook.Sheets("Sheet1").Cells(1, 2).Value = "description"
    RelistBook.Sheets("Sheet1").Cells(1, 3).Value = "ca actual"
    RelistBook.Sheets("Sheet1").Cells(1, 4).Value = "inline"
    
    
    relistLastRow = RelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
    
    
    'return to proper worksheet
    relistSheet.Activate
    
    Range("A2:M" & lastRow).Sort key1:=Range("A2:A" & lastRow), _
      order1:=xlAscending, Header:=xlYes
    
    For i = 2 To lastRow
        If Cells(i, 8).Value > 0 Then
            If Cells(i, 13).Value > 0 Then
                Cells(i, 1).Copy Destination:=RelistBook.Sheets("Sheet1").Range("A" & relistLastRow)
                Cells(i, 2).Copy Destination:=RelistBook.Sheets("Sheet1").Range("B" & relistLastRow)
                Cells(i, 13).Copy Destination:=RelistBook.Sheets("Sheet1").Range("C" & relistLastRow)
                Cells(i, 10).Copy Destination:=RelistBook.Sheets("Sheet1").Range("D" & relistLastRow)
                relistLastRow = RelistBook.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).row
            End If
        End If
    Next i
    
    With RelistBook.Sheets("Sheet1").UsedRange
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 8
        .Interior.ColorIndex = xlNone
    End With

    RelistBook.Sheets("Sheet1").Columns("A:D").AutoFit
    
    Dim RelistEmail As Range
    Set RelistEmail = RelistBook.Sheets("Sheet1").Range("A1:D" & relistLastRow)
    
    Call report_emails.EmailReport("email-list", "", "Inventory Relist ", RangetoHTML(RelistEmail), "")
    
    RelistBook.Close SaveChanges:=False

End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
