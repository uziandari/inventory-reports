Sub GatherReports(User_ID, Password)

    Dim InventoryArray As Variant
    Dim ReportArray As Variant
    Dim jDailySheet As Worksheet

    Dim compUserName As String
    compUserName = CreateObject("WScript.Network").Username
    
    'Separates reports to email and inventory reports
    ReportArray = Array("dupe_loc", "dupe_upc", "free_locations", "wholesale_committed", "negative_ns", "moves", "ca_greater_ns", "na_with_quantity", "dropship_incorrects", "backstock_pulls", "b2b_quantity")
    InventoryArray = Array("less_nine", "alerts", "delist", "relist", "relist_pushed", "jdaily_quantity")
    
    Application.ScreenUpdating = False
    
    'Sets Negatives to 0
    'Emails negative
    Call NegativePullFromDB(User_ID, Password)

    Call PullFromDB(ReportArray, "Reports ", compUserName)
    'Email Reports File
    '
    Call report_emails.EmailReport("email", "email", "Daily Inventory " & Format(CStr(Now), "yyyy_mm_dd"), "Attached.", "C:\Users\" & compUserName & "\Desktop\Reports " & Format(CStr(Now), "yyyy_mm_dd") & ".xlsx")
    '
    'End Email Reports File
    
    Call PullFromDB(InventoryArray, "InventoryReports ", compUserName)
    
    'call relist push to create report
    Call report_calculations.calculateRelistToPush
    'end call relist
    
    Set jDailySheet = Workbooks("InventoryReports " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("jdaily_quantity")
    
    jDailySheet.Activate
    
    Call SaveFileToImport("jdaily_quantity", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")
    
    Call report_calculations.calculateB2bInventory

    Call report_emails.EmailReport("email", "", "B2B_Upload " & Format(CStr(Now), "yyyy_mm_dd"), "Attached.", "C:\Users\" & compUserName & "\Desktop\b2bUpload.csv")

    Application.DisplayAlerts = False
    Workbooks("Reports " & Format(CStr(Now), "yyyy_mm_dd")).Close
    Workbooks("jdaily_quantity").Close
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    User_ID = ""
    Password = ""
    
End Sub


Sub PullFromDB(ReportArray, saveFile, compUserName)

    Dim report As Variant
    
    Set NewBook = Workbooks.Add
    
    Worksheets("sheet1").Name = "Special"
    
    For Each report In ReportArray
        
        With NewBook
            
            ' Adds sheet to notebook and names sheet after report
            
            NewBook.Sheets.Add(Before:=.Sheets(.Sheets.Count)).Name = report
            Worksheets(report).Activate

            Dim sqlQuery As String
            
            If Worksheets(report).Name = "free_locations" Then
                'Call FreeLocations
                sqlQuery = "SELECT * FROM `inventory`.`free_locations` ORDER BY 1"
            ElseIf Worksheets(report).Name = "negative_ns" Then
                'Call NegativeStock
                sqlQuery = "SELECT `sku`, `description`, `location`, `stock`, `inline`, `purchase_price` FROM `inventory`.`ns_inventory` WHERE stock < 0"
            ElseIf Worksheets(report).Name = "moves" Then
                'Call PotentialMoves
                 sqlQuery = "SELECT `main_inventory`.`sku`, `main_inventory`.`ns_description`, `main_inventory`.`stock` As quantity, `main_inventory`.`committed`, " _
                        & "`main_inventory`.`upc`, `main_inventory`.`location`, `main_inventory`.`backstock`, `location_table`.`bin_size`, '' As newLoc " _
                        & "FROM `inventory`.`main_inventory` " _
                        & "JOIN `invrec`.`location_table` ON `main_inventory`.`location` = `location_table`.`location` " _
                        & "WHERE (stock < 11 AND `location_table`.`bin_size` IN (" & Chr(34) & "LG" & Chr(34) & "," & Chr(34) & "BULK" & Chr(34) & ")) " _
                        & "OR (stock < 5 AND `location_table`.bin_size IN (" & Chr(34) & "MD" & Chr(34) & "," & Chr(34) & "CL" & Chr(34) & ")) " _
                        & "OR (stock = 1 AND `location_table`.bin_size IN (" & Chr(34) & "SM" & Chr(34) & "," & Chr(34) & "SHOE" & Chr(34) & ")) " _
                        & "ORDER BY `main_inventory`.`location`"
             ElseIf Worksheets(report).Name = "ca_greater_ns" Then
                sqlQuery = "SELECT `main_inventory`.`sku`, `description`, `total` As CA_total, `location`, `upc`, `stock` As NS_Quantity, `committed`, `inline`, `receipt` FROM `inventory`.`main_inventory` " _
                        & "LEFT JOIN `inventory`.`dropship` ON `inventory`.`main_inventory`.`sku` = `inventory`.`dropship`.`sku` WHERE `inventory`.`dropship`.`sku` IS NULL " _
                        & "AND `main_inventory`.`total` > `main_inventory`.`stock` AND `main_inventory`.`location` NOT LIKE " & Chr(34) & "%/DROPSHIP" & Chr(34)
            ElseIf Worksheets(report).Name = "na_with_quantity" Then
                sqlQuery = "SELECT `main_inventory`.`sku`, `description`, `total` As CA_total, `location`, `upc`, `stock` As NS_Quantity, `committed`, `inline`, `main_inventory`.`receipt`, " _
                        & "`purchase_price` * `stock` * -1 As Total_Cost FROM `inventory`.`main_inventory` " _
                        & "WHERE `main_inventory`.`bucket` = 0 AND `location` IN (" & Chr(34) & "NA" & Chr(34) & ", " & Chr(34) & "DROPSHIP" & Chr(34) & ") " _
                        & "AND `sku` NOT LIKE " & Chr(34) & "Ball%" & Chr(34) & " AND `flag` NOT LIKE " & Chr(34) & "%recount%" & Chr(34) _
                        & " AND `stock` > 0 "
            ElseIf Worksheets(report).Name = "dropship_incorrects" Then
                'Call Dropship
                sqlQuery = "SELECT `main_inventory`.`sku`, `description`, `total`, `available`, `location`, `upc`, `stock`, `committed`, `inline`, `dropship`.`sku` AS isDropship FROM `inventory`.`main_inventory` " _
                        & "LEFT JOIN `inventory`.`dropship` ON `inventory`.`main_inventory`.`sku` = `dropship`.`sku` " _
                        & "WHERE (`inventory`.`dropship`.`sku` IS NULL " _
                        & "AND `location` LIKE " & Chr(34) & "DROPSHIP" & Chr(34) & "AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "90%" & Chr(34) _
                        & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "Ball%" & Chr(34) & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "misc%" & Chr(34) & ") " _
                        & "OR (`main_inventory`.`available` > 9900 AND `inventory`.`dropship`.`sku` IS NULL AND `location` NOT LIKE " & Chr(34) & "%/DROPSHIP" & Chr(34) _
                        & " AND `location` NOT LIKE " & Chr(34) & "OFFICE" & Chr(34) & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "90%" & Chr(34) _
                        & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "Ball%" & Chr(34) & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "misc%" & Chr(34) & ") " _
                        & "OR (`main_inventory`.`available` <> 0 AND `inventory`.`dropship`.`sku` IS NOT NULL AND `location` NOT LIKE " & Chr(34) & "NA" & Chr(34) _
                        & " AND `location` NOT LIKE " & Chr(34) & "DROPSHIP" & Chr(34) & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "90%" & Chr(34) _
                        & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "Ball%" & Chr(34) & " AND `main_inventory`.`sku` NOT LIKE " & Chr(34) & "misc%" & Chr(34) & ")"
            ElseIf Worksheets(report).Name = "backstock_pulls" Then
                sqlQuery = "SELECT `daily_sales`.`sku`, `main_inventory`.`ns_description`, `main_inventory`.`location`, `main_inventory`.`backstock`, `main_inventory`.`upc`, " _
                        & "`main_inventory`.`stock`, `main_inventory`.`committed`, `main_inventory`.`pending_shipment` FROM `inventory`.`daily_sales` " _
                        & "LEFT JOIN `inventory`.`main_inventory` ON `daily_sales`.`sku` = `main_inventory`.`sku` WHERE `committed` > 0 AND " _
                        & "`backstock` NOT IN (" & Chr(34) & "NA" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ") " _
                        & "GROUP BY `main_inventory`.`sku` ORDER BY `pending_shipment` DESC" 
            Else
                sqlQuery = "SELECT " & reportSelect(report) & " FROM `inventory`.`main_inventory` WHERE `inventory`.`main_inventory`." & report & " = 1"
            End If

                'pulls relevant info from DB
                With ActiveSheet.ListObjects.Add(SourceType:=0, Source:="ODBC;DSN=localDB;" _
                    , Destination:=Range("$A$1")).QueryTable
                    .CommandType = xlCmdSql
                    .CommandText = sqlQuery
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .BackgroundQuery = True
                    .RefreshStyle = xlInsertDeleteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .PreserveColumnInfo = True
                    .SourceConnectionFile = _
                    "C:\Users\" & compUserName & "\OneDrive\Documents\My Data Sources\local_db_connection.odc"
                    .ListObject.DisplayName = "Table_" & report
                    .Refresh BackgroundQuery:=False
                
                End With
            
            FormatTable
            
        End With
    Next
    
    Application.ScreenUpdating = True
    
    Worksheets("Special").Activate
    
    'Save File
    Call SaveFileToImport(saveFile & Format(CStr(Now), "yyyy_mm_dd"), 51, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")
    
End Sub

Function reportSelect(report)
            
'Changes the select statement based on which report is queried
            
    If report = "less_nine" Or report = "alerts" Then
        reportSelect = "sku, description, total, available, pending_checkout + pending_payment As pending, GREATEST(pending_shipment, committed) As committed, location, backstock, upc, stock, '' As actual, '' As ADJUSTMENT, inline, flag, receipt"
    ElseIf report = "dupe_loc" Or report = "dupe_upc" Then
        reportSelect = "sku, description, location, upc"
    ElseIf report = "relist" Then
        reportSelect = "sku, description, location, backstock, upc, stock, committed, '' As actual, '' As newLoc, inline, flag, receipt"
    ElseIf report = "relist_pushed" Then
        reportSelect = "sku, description, total, available, pending_checkout, pending_payment, pending_shipment, committed, location, backstock, upc, stock, inline, flag, receipt"
    ElseIf report = "delist" Then
        reportSelect = "sku, description, stock - committed As NSQty, inline"
    ElseIf report = "jdaily_quantity" Then
        reportSelect = "sku As `Product SKU`, GREATEST(available - 2, 0) As `Stock Level`"
    ElseIf report = "b2b_quantity" Then
        reportSelect = "sku, GREATEST(available, 0) As `available`"    
    Else
        reportSelect = "sku, description, total, available, pending_checkout + pending_payment As pending, GREATEST(pending_shipment, committed) As committed, location, backstock, upc, stock, inline, flag, receipt"
    End If
End Function

Sub FormatTable()

    With ActiveSheet.UsedRange
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 8
    End With
    
End Sub

Sub NegativePullFromDB(User_ID, Password)

    Set NewBook = Workbooks.Add
    
    Worksheets("sheet1").Name = "NegativeInventory"
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:="ODBC;DSN=localDB;" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT sku, description, total, available, pending_checkout, pending_payment, pending_shipment, committed, location, upc, stock, inline, flag, receipt FROM `inventory`.`main_inventory` WHERE `main_inventory`.available < 0"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceConnectionFile = _
        "C:\Users\" & compUserName & "\OneDrive\Documents\My Data Sources\local_db_connection.odc"
        .ListObject.DisplayName = "Table_negative"
        .Refresh BackgroundQuery:=False

    End With
    
    'Saves the file
    Call SaveFileToImport("negative_inventory " & Format(CStr(Now), "yyyy_mm_dd"), 51, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")

    'Exits sub if there is no negative inventory
    Set negativeSheet = Workbooks("negative_inventory " & Format(CStr(Now), "yyyy_mm_dd")).Sheets("NegativeInventory")
  
    Dim lastRow As Long
    lastRow = negativeSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    If lastRow = 1 Then
        Exit Sub
    End If
    'End exit code
    
    Call UpdateNegative(negativeSheet, lastRow, User_ID, Password)
    
    
    'email list
    Call report_emails.EmailReport("email", "", "Negative Inventory " & Format(CStr(Now), "yyyy_mm_dd"), "Attached.", CreateObject("WScript.Shell").specialfolders("Desktop") & "\negative_inventory " & Format(CStr(Now), "yyyy_mm_dd") & ".xlsx")

    Workbooks("negative_inventory " & Format(CStr(Now), "yyyy_mm_dd")).Close

End Sub

Sub UpdateNegative(negativeSheet, lastRow, User_ID, Password)

    Set adjustmentBook = Workbooks.Add
    Call upload_files.caUploadFileInitialize(adjustmentBook)

    'return to proper worksheet
    negativeSheet.Activate

    ' loop through input
    For i = 2 To lastRow
        Dim sku As String, adjustmentType As String, flag As Variant, flagDescription As Variant, adjustment As Integer
    
        sku = Cells(i, 1).Value
        adjustmentType = "Absolute"
        adjustment = 0
        flag = Null
        flagDescription = Null
        
        Call UpdateDBNegatives(sku, User_ID, Password)

        Call upload_files.caUploadFileAdd(adjustmentBook, sku, adjustmentType, adjustment, flag, flagDescription)

    Next
    
    adjustmentBook.Activate
    'Save File
    Call SaveFileToImport("negativeCA", 6, CreateObject("WScript.Shell").specialfolders("Desktop") & "\")
    
    Workbooks("negativeCA").Close

End Sub

Sub UpdateDBNegatives(sku, User_ID, Password)

    Dim SQLStr As String
    Dim Server_Name As String
    Dim Database_Name As String
    
    Set rs = CreateObject("ADODB.Recordset")
    Server_Name = "localhost"
    Database_Name = "inventory"
    User_ID = User_ID
    Password = Password
    Port = "3306"
    
    SQLStr = "UPDATE inventory.main_inventory SET total = total - available, available = 0 WHERE main_inventory.sku LIKE " & "'" & sku & "'"
    
    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
            Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
            ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    rs.Open SQLStr, Cn, adOpenStatic

End Sub

Sub DailyAdjustments()

    Dim compUserName As String
    compUserName = CreateObject("WScript.Network").Username

    Set NewBook = Workbooks.Add
    
    Worksheets("sheet1").Name = "adjustments"
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:="ODBC;DSN=localDB;" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM `invrec`.`daily_adjustment`"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceConnectionFile = _
        "C:\Users\" & compUserName & "\OneDrive\Documents\My Data Sources\local_db_connection.odc"
        .ListObject.DisplayName = "Table_adjustments"
        .Refresh BackgroundQuery:=False

    End With
    
    'Saves the file
    Call SaveFileToImport("inventory_adjustment_upload " & Format(CStr(Now), "yyyy_mm_dd"), 6, "C:\Users\" & compUserName & "\Desktop\adjustments\")
     
    'emails adjustments
    Call report_emails.EmailReport("email", "", "NS Adjustments To Upload ", "Attached.", "C:\Users\" & compUserName & "\Desktop\adjustments\inventory_adjustment_upload " & Format(CStr(Now), "yyyy_mm_dd") & ".csv")

    'calls routine to email high value adjustments
    Call report_emails.EmailHighValueAdjustments

End Sub








