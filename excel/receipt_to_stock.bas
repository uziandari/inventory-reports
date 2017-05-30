Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub populate_receipts()

  skuArray = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "30", "31", "32", "33")
  bulkArray = Array("02", "03", "07", "09")
  ballArray = Array("01", "17")
  Dim SQLStr As String
  
  Set rs = CreateObject("ADODB.Recordset")
  Server_Name = "localhost"
  Database_Name = "inventory"
  Port = "3306"
  'user, pw go here
  User_ID = "root"
  Password = "woompa1"

  For Each sku In skuArray

    Dim binSize As String, qty As Integer

    If IsInArray(sku, bulkArray) Then
      binSize = Array("BULK", "BULKLQ")
      qty = 4
    Else If IsInArray(sku, ballArray)
      GoTo NextIteration
    End If
    
    OpenStr = "UPDATE inventory.receipt_to_stock SET bin_size = '" & binSize(0) & "' WHERE sku LIKE '" & sku & "%' AND old_bin IN ('', 'NA', 'DROPSHIP') AND qty_received > " & qty
    SafeOff = "SET SQL_SAFE_UPDATES=0"
    SQLStr = "UPDATE inventory.receipt_to_stock LEFT JOIN fl ON receipt_to_stock.bin_size = fl.bin_size " _
            & "SET new_bin = (SELECT location FROM fl WHERE bin_size = receipt_to_stock.bin_size ORDER BY RAND() LIMIT 1)"
    SafeOn = "SET SQL_SAFE_UPDATES=1"

    Set Cn = CreateObject("ADODB.Connection")
    Cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & _
    Server_Name & ";Port=" & Port & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open OpenStr, Cn, adOpenStatic
    rs.Open SafeOff, Cn, adOpenStatic
    rs.Open SQLStr, Cn, adOpenStatic
    rs.Open SafeOn, Cn, adOpenStatic

NextIteration:
  Next


End Sub
