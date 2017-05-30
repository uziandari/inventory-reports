Sub initializeMain(User_ID, Password)

  'PULL IN USER, PW AS ARGUMENTS

  Dim SQLStr As String
  Dim Server_Name As String
  Dim Database_Name As String

  Dim currentDay As String
  currentDay = Format(Date, "m/d")

  Set rs = CreateObject("ADODB.Recordset")
  Server_Name = "localhost"
  Database_Name = "inventory"
  Port = "3306"
  'user, pw go here
  User_ID = User_ID
  Password = Password


  Set queryDict = CreateObject("Scripting.Dictionary")

  Dim naLoc As String, dsLoc As String, inlinefield As String, recountFlag As String, absoluteFlag As String
  naLoc = "%NA%"
  dsLoc = "Dropship%"
  inlinefield = "Yes"
  recountFlag = "%recount%"
  absoluteFlag = "absolute final%"

  queryDict.Add "clearMainInventory", "TRUNCATE inventory.main_inventory"
  queryDict.Add "loadMainInventory", "INSERT INTO inventory.main_inventory (`sku`, `description`, `total`, `available`, `pending_checkout`, `pending_payment`, `pending_shipment`, `flag`, `parent_sku`, " _
                & "`label`, `ns_description`, `location`, `backstock`, `upc`, `stock`, `committed`, `head_cover`, `inline`, `purchase_price`, `receipt`, `is_blocked`, `img_url`) " _
                & "SELECT `inventory`.`ca_inventory`.`sku`, `inventory`.`ca_inventory`.`description`, `inventory`.`ca_inventory`.`total`, `inventory`.`ca_inventory`.`available`, " _
                & "`inventory`.`ca_inventory`.`pending_checkout`, `inventory`.`ca_inventory`.`pending_payment`, `inventory`.`ca_inventory`.`pending_shipment`, `inventory`.`ca_inventory`.`flag`, " _
                & "`inventory`.`ca_inventory`.`parent_sku`, `inventory`.`ca_inventory`.`label`, `inventory`.`ns_inventory`.`description`, `inventory`.`ns_inventory`.`location`, " _
                & "`inventory`.`ns_inventory`.`backstock`, `inventory`.`ns_inventory`.`upc`, `inventory`.`ns_inventory`.`stock`, `inventory`.`ns_inventory`.`committed`, " _
                & "`inventory`.`ns_inventory`.`head_cover`, `inventory`.`ns_inventory`.`inline`, `inventory`.`item_cost`.`average_cost`, `inventory`.`receipt_date`.`receipt`, " _
                & "`inventory`.`ca_inventory`.`blocked`, `inventory`.`ca_inventory`.`img_url` " _
                & "FROM `inventory`.`ca_inventory` INNER JOIN `inventory`.`ns_inventory` ON `inventory`.`ca_inventory`.`sku` = `inventory`.`ns_inventory`.`sku` " _
                & "LEFT JOIN `inventory`.`receipt_date` ON `inventory`.`ca_inventory`.`sku` = `inventory`.`receipt_date`.`sku` " _
                & "LEFT JOIN `inventory`.`item_cost` ON `inventory`.`ca_inventory`.`sku` = `inventory`.`item_cost`.`sku`"
  queryDict.Add "removeSafeUpdates", "SET SQL_SAFE_UPDATES=0"
  queryDict.Add "updateNaLocation", "UPDATE inventory.main_inventory SET inventory.main_inventory.location = " & Chr(34) & "NA" & Chr(34) & " WHERE inventory.main_inventory.location = " & Chr(34) & Chr(34)
  queryDict.Add "updateInline", "UPDATE inventory.main_inventory SET inventory.main_inventory.inline = " & Chr(34) & "No" & Chr(34) & " WHERE inventory.main_inventory.inline = " & Chr(34) & Chr(34)
  queryDict.Add "updateBucket", "UPDATE inventory.main_inventory JOIN inventory.bucket ON inventory.main_inventory.sku = inventory.bucket.sku SET inventory.main_inventory.bucket = 1 WHERE inventory.bucket.sku IS NOT NULL"
  queryDict.Add "updateJohnDailyQuantity", "UPDATE inventory.main_inventory JOIN inventory.big_commerce_daily ON inventory.main_inventory.sku = inventory.big_commerce_daily.sku SET inventory.main_inventory.jdaily_quantity = 1 WHERE inventory.big_commerce_daily.sku IS NOT NULL"
  queryDict.Add "updateB2bDailyQuantity", "UPDATE inventory.main_inventory JOIN inventory.b2b_daily ON inventory.main_inventory.sku = inventory.b2b_daily.sku SET inventory.main_inventory.b2b_quantity = 1 WHERE inventory.b2b_daily.sku IS NOT NULL"
  queryDict.Add "updateRemoves", "UPDATE inventory.main_inventory JOIN inventory.removes ON inventory.main_inventory.sku = inventory.removes.sku SET inventory.main_inventory.remove = 1 WHERE inventory.removes.sku IS NOT NULL"
  queryDict.Add "updateRelistFromFile", "UPDATE inventory.main_inventory INNER JOIN inventory.relists ON inventory.main_inventory.sku = inventory.relists.sku SET inventory.main_inventory.relist = 1 WHERE inventory.relists.sku IS NOT NULL"
  queryDict.Add "updateRelistFromFlag", "UPDATE inventory.main_inventory SET inventory.main_inventory.relist = 1 WHERE inventory.main_inventory.flag LIKE " _
                & Chr(34) & "final recount " & currentDay & "%" & Chr(34)
  queryDict.Add "updateDailyRemoves", "UPDATE inventory.main_inventory INNER JOIN inventory.removes_daily ON inventory.main_inventory.sku = inventory.removes_daily.sku SET inventory.main_inventory.remove = 1 WHERE inventory.removes_daily.sku IS NOT NULL"
  queryDict.Add "updateLessThanNine", "UPDATE inventory.main_inventory SET inventory.main_inventory.less_nine = 1 WHERE (available <= 9 and available > 0) " _
                & "AND (flag NOT LIKE " & Chr(34) & "%final%" & Chr(34) & " AND flag NOT LIKE " & Chr(34) & "%Inline%" & Chr(34) & ") " _
                & "AND (location <> " & Chr(34) & "NA" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "%Dropship%" & Chr(34) & ") " _
                & "AND bucket = 0 AND sku NOT LIKE " & Chr(34) & "9%" & Chr(34)
  queryDict.Add "updateDelist", "UPDATE inventory.main_inventory SET inventory.main_inventory.delist = 1 WHERE sku NOT LIKE " & Chr(34) & "9%" & Chr(34) _
                & " AND stock = committed AND location NOT LIKE " & Chr(34) & naLoc & Chr(34) _
                & " AND location NOT LIKE " & Chr(34) & "%DROPSHIP" & Chr(34) _
                & " AND inventory.main_inventory.flag NOT LIKE " & Chr(34) & absoluteFlag & Chr(34) _
                & " AND inventory.main_inventory.flag NOT LIKE " & Chr(34) & recountFlag & Chr(34) & " AND relist = 0 AND remove = 0 "
  queryDict.Add "updateAlerts", "UPDATE inventory.main_inventory SET inventory.main_inventory.alerts = 1 WHERE (available = 0 AND sku NOT LIKE " & Chr(34) & "9%" & Chr(34) & " AND flag NOT LIKE " & Chr(34) & absoluteFlag & Chr(34) _
                & " AND flag NOT LIKE " & Chr(34) & "%ebay%" & Chr(34) & " AND flag NOT LIKE " & Chr(34) & "%hold%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & naLoc & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & dsLoc & Chr(34) _
                & " AND flag NOT LIKE " & Chr(34) & recountFlag & Chr(34) & " AND less_nine = 0 AND delist = 0 AND relist = 0 AND remove = 0 AND bucket = 0)"
  queryDict.Add "pendingFromAlerts", "UPDATE inventory.main_inventory SET inventory.main_inventory.alerts = 0 " _
                & "WHERE inventory.main_inventory.alerts = 1 AND inventory.main_inventory.stock > 0 AND (inventory.main_inventory.pending_checkout + inventory.main_inventory.pending_payment = inventory.main_inventory.stock)"
  queryDict.Add "safetyFromAlerts", "UPDATE inventory.main_inventory SET inventory.main_inventory.alerts = 0 WHERE inventory.main_inventory.alerts = 1 " _
                & "AND inventory.main_inventory.stock > 0 AND (GREATEST(inventory.main_inventory.pending_shipment, inventory.main_inventory.committed) > 0) AND (GREATEST(inventory.main_inventory.pending_shipment, inventory.main_inventory.committed) + 2 >= inventory.main_inventory.stock)"
  queryDict.Add "updateRelistPushed", "UPDATE inventory.main_inventory SET inventory.main_inventory.relist_pushed = 1, inventory.main_inventory.relist = 0 " _
                & "WHERE (inventory.main_inventory.relist = 1 AND (inventory.main_inventory.total <> 0 OR inventory.main_inventory.`committed` <> 0 OR inventory.main_inventory.bucket = 1))"
  queryDict.Add "findDuplicateLocations", "UPDATE inventory.main_inventory INNER JOIN (SELECT inventory.main_inventory.location FROM inventory.main_inventory WHERE inventory.main_inventory.stock> 0 AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "C%" & Chr(34) _
                & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "Z CAB%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "L%" & Chr(34) & "AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "M%" & Chr(34) _
                & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "N%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "O%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "I24%" & Chr(34) _
                & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "I26%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & dsLoc & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "GNC%" & Chr(34) _
                & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "J%" & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & "K%" & Chr(34) _
                & " GROUP BY inventory.main_inventory.location HAVING COUNT(inventory.main_inventory.location) > 1) AS secondary_inventory ON inventory.main_inventory.location = secondary_inventory.location SET inventory.main_inventory.dupe_loc = 1"
  queryDict.Add "findDuplicateUpcs", "UPDATE inventory.main_inventory INNER JOIN (SELECT inventory.main_inventory.upc FROM inventory.main_inventory " _
                & "WHERE inventory.main_inventory.upc NOT LIKE " & Chr(34) & Chr(34) & " AND inventory.main_inventory.upc NOT LIKE  " & Chr(34) & "0000000000%%" & Chr(34) _
                & " GROUP BY inventory.main_inventory.upc HAVING COUNT(inventory.main_inventory.upc) > 1) AS secondary_inventory " _
                & "ON inventory.main_inventory.upc = secondary_inventory.upc SET inventory.main_inventory.dupe_upc = 1"
  queryDict.Add "capitalizeLocations", "UPDATE inventory.main_inventory SET inventory.main_inventory.location = UPPER( inventory.main_inventory.`location` )"
  queryDict.Add "updateBlankLocatoins", "UPDATE inventory.main_inventory SET inventory.main_inventory.location = " & Chr(34) & "NA" & Chr(34) & " WHERE inventory.main_inventory.location LIKE " & Chr(34) & Chr(34)
  queryDict.Add "findWholesale", "UPDATE inventory.main_inventory INNER JOIN inventory.wholesale_pending ON inventory.main_inventory.sku = inventory.wholesale_pending.sku SET inventory.main_inventory.wholesale_committed = 1 " _
                & "WHERE inventory.main_inventory.available = 0" _
                & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & naLoc & Chr(34) & " AND inventory.main_inventory.location NOT LIKE " & Chr(34) & dsLoc & Chr(34) _
                & "AND inventory.main_inventory.`committed` > 0"
  queryDict.Add "clearLocationsTable", "TRUNCATE inventory.free_locations"
  queryDict.Add "findLocations", "INSERT INTO inventory.free_locations (location, bin_size) SELECT location, bin_size FROM invrec.location_table WHERE location NOT IN (SELECT location FROM inventory.ns_inventory) " _
                & "AND location NOT IN (SELECT location FROM inventory.location_removes)"
  queryDict.Add "applySafeUpdates", "SET SQL_SAFE_UPDATES=1"


  For Each query In queryDict.Keys

    SQLStr = queryDict.Item(query)


    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
    Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open SQLStr, Cn, adOpenStatic

  Next

  Set queryDict = Nothing

End Sub
