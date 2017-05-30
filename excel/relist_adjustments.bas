Sub AddAdjustmentsToDB()

    Dim sqlStr As String, sqlTruncateStr As String, sqlUpdateSoftGoodsStr As String, sqlUpdateHardGoodsStr As String
    Dim Server_Name As String
    Dim Database_Name As String
    Dim compUserName As String
    Dim softGoodsNumber As Integer, hardGoodsNumber As Integer
    
    compUserName = CreateObject("WScript.Network").Username
    
    Dim lngEoM As Long
    lngEoM = Date - Day(Date)
     
    If Date = lngEoM + Choose(Weekday(lngEoM, vbMonday), 0, 0, 0, 0, 0, 2, 1) Then
        MsgBox "It's the first day of the month --update your report numbers."
    End If
 
    
    'need to update this at the beginning of the month!!!!!!
    softGoodsNumber = 2578
    hardGoodsNumber = 2579
    '-need to update this at the beginning of the month!!!!!!-
    
    Set rs = CreateObject("ADODB.Recordset")
    Server_Name = "localhost"
    Database_Name = "invrec"
    Port = "3306"
    
    Login.Show
    If Len(Trim(User_ID)) = 0 Then
      User_ID = "root"
    End If
    If Len(Trim(Password)) = 0 Then
      MsgBox ("You didn't provide a password.")
      Exit Sub
    End If

    If Format(FileDateTime("C:\Users\" & compUserName & "\Desktop\adjustments\local_db_daily_adjustments.csv"), "mm/dd/yyyy") <> Format(Date, "mm/dd/yyyy") Then
      MsgBox ("You need to update local_db_daily_adjustments.csv")
      Environment.Exit (0)
    End If
    
    sqlTruncateStr = "TRUNCATE `invrec`.`daily_adjustment`"
    sqlStr = "LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\" & compUserName _
    & "\\Desktop\\adjustments\\local_db_daily_adjustments.csv' REPLACE INTO TABLE `invrec`.`daily_adjustment` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '" _
    & Chr(34) & "' ESCAPED BY '" & Chr(34) & "' LINES TERMINATED " _
    & "BY '\r\n' IGNORE 1 LINES (`SKU`, `QTY`)"
    sqlUpdateSoftGoodsStr = "UPDATE invrec.daily_adjustment SET invrec.daily_adjustment.`Inventory Adjustment #` = 2578 " _
                            & "WHERE LEFT(`invrec`.`daily_adjustment`.`SKU`, 2) IN ('01', '02', '03', '04', '05', '06', '16', '17', '18', '19', '20') " _
                            & "OR LEFT(`invrec`.`daily_adjustment`.`SKU`, 1) = '3'"
                            
    sqlUpdateHardGoodsStr = "UPDATE invrec.daily_adjustment SET invrec.daily_adjustment.`Inventory Adjustment #` = 2579 " _
                            & "WHERE LEFT(`invrec`.`daily_adjustment`.`SKU`, 2) IN ('07', '08', '09', '10', '11', '12', '13', '14', '15', '21') " _
                            & "OR LEFT(`invrec`.`daily_adjustment`.`SKU`, 1) = '6'"

    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
            Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
            ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    rs.Open sqlTruncateStr, Cn, adOpenStatic
    rs.Open sqlStr, Cn, adOpenStatic
    rs.Open sqlUpdateSoftGoodsStr, Cn, adOpenStatic
    rs.Open sqlUpdateHardGoodsStr, Cn, adOpenStatic



    User_ID = ""
    Password = ""

End Sub

