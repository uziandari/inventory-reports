Public User_ID As String
Public Password As String


Sub truncateTables()

  tableArray = Array("bucket", "big_commerce_daily", "ca_inventory", "relists", "main_inventory", "ns_inventory", "receipt_date", "removes_daily", "wholesale_pending", "dropship", "adjustment", "item_cost", "daily_sales", "previous_locations", "previous_upcs", "previous_receipts")

  Dim SQLStr As String
  Dim Server_Name As String
  Dim Database_Name As String


  Set rs = CreateObject("ADODB.Recordset")
  Server_Name = "localhost"
  Database_Name = "inventory"
  Port = "3306"
  
  Login.Show
  If Len(Trim(User_ID)) = 0 Then
    User_ID = "root"
  End If
  If Len(Trim(Password)) = 0 Then
    MsgBox ("You didn't provide a password.")
    Exit Sub
  End If

  For Each table In tableArray

    SQLStr = "TRUNCATE `inventory`.`" & table & "`"


    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
    Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open SQLStr, Cn, adOpenStatic

  Next

  Call loadCsv(User_ID, Password)
  Call filter_sql.initializeMain(User_ID, Password)
  Call db_format.GatherReports(User_ID, Password)

  User_ID = ""
  Password = ""

End Sub

Sub loadCsv(user, Password)

  'declare and set up dictionary
  Set csvDict = CreateObject("Scripting.Dictionary")

  'add values to dictionary
  csvDict.Add "daily_bucket", "bucket"
  csvDict.Add "daily_bc", "big_commerce_daily"
  csvDict.Add "ca_import", "ca_inventory"
  csvDict.Add "relist", "relists"
  csvDict.Add "ns_import", "ns_inventory"
  csvDict.Add "receipt_date_import", "receipt_date"
  csvDict.Add "prevdelists", "removes_daily"
  csvDict.Add "wholesale_committed", "wholesale_pending"
  csvDict.Add "dropship_skus", "dropship"
  csvDict.Add "adj_import", "adjustment"
  csvDict.Add "item_cost_import", "item_cost"
  csvDict.Add "daily_sales_import", "daily_sales"
  csvDict.Add "prev_location_import", "previous_locations"
  csvDict.Add "prev_upc_import", "previous_upcs"
  csvDict.Add "prev_receipt_import", "previous_receipts"
  csvDict.Add "daily_b2b", "b2b_daily"



  Set rs = CreateObject("ADODB.Recordset")
  Server_Name = "localhost"
  Database_Name = "inventory"
  Port = "3306"
  User_ID = user
  Password = Password

  Dim compUserName As String
  compUserName = CreateObject("WScript.Network").Username
  
  For Each table In csvDict.Keys

    Dim inj As String

    If Format(FileDateTime("C:\Users\" & compUserName & "\Desktop\database_files\inventory\import_files\" & table & ".csv"), "mm/dd/yyyy") <> Format(Date, "mm/dd/yyyy") Then
      MsgBox ("You need to update " & table & ".csv")
      Environment.Exit (0)
    End If

    If table = "ca_import" Then
      inj = "`sku`, `description`, `total`, `available`, `pending_checkout`, `pending_payment`, `pending_shipment`, `flag`, `blocked`, `parent_sku`, `label`,`img_url`"
    ElseIf table = "ns_import" Then
      inj = "`sku`, `description`, `location`, `backstock`, `upc`, `stock`, `committed`, `head_cover`, `inline`, `purchase_price`"
    ElseIf table = "receipt_date_import" Then
      inj = "`sku`, `receipt`"
    ElseIf table = "adj_import" Then
      inj = "`sku`, `adjustment_date`, `adjustment_amount`"
    ElseIf table = "item_cost_import" Then
      inj = "`sku`, `average_cost`, `purchase_price`"
    ElseIf table = "daily_sales_import" Then
      inj = "`order_date`, `customer`, `order_id`, `sku`"
    ElseIf table = "prev_location_import" Then
      inj = "`sku`, `changeDate`, `changeField`, `changeLoc`"
    ElseIf table = "prev_upc_import" Then
      inj = "`sku`, `changeDate`, `changeUpc`"
    ElseIf table = "prev_receipt_import" Then
      inj = "`document`, `receiptDate`, `sku`, `quantity`, `type`"
    Else
      inj = "`sku`"
    End If

    SQLStr = "LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\" & compUserName _
    & "\\Desktop\\database_files\\inventory\\import_files\\" _
    & table & ".csv" & "' REPLACE INTO TABLE `inventory`.`" & csvDict.Item(table) & "` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '" _
    & Chr(34) & "' ESCAPED BY '" & Chr(34) & "' LINES TERMINATED " _
    & "BY '\r\n' IGNORE 1 LINES (" & inj & ")"

  


    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
    Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open SQLStr, Cn, adOpenStatic


  Next

  'reset dictionary
  Set csvDict = Nothing

End Sub



