Option Explicit

Private Sub btnAdd_Click()
    If Me.txtTrackingNumber.Value = "" Then
        MsgBox "Please enter a tracking number.", vbExclamation, "Item Return"
        Me.txtTrackingNumber.SetFocus
        Exit Sub
    End If
    
    If Me.txtOrderNumber.Value = "" Then
        MsgBox "Please enter an order number.", vbExclamation, "Item Return"
        Me.txtOrderNumber.SetFocus
        Exit Sub
    End If
    
    If Me.txtUpc.Value = "" Then
        MsgBox "Please enter the item's UPC.", vbExclamation, "Item Return"
        Me.txtUpc.SetFocus
        Exit Sub
    End If
    
    
    'add data to worksheet
    Dim returnsFormSheet As Worksheet
    Set returnsFormSheet = Workbooks("Returns").Sheets("returns")
  
    Dim lastrow As Long
    lastrow = returnsFormSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    With returnsFormSheet.Range("A1")
     .Offset(lastrow, 0).Value = Format(Now(), "MMM-DD-YYYY")
     .Offset(lastrow, 1).Value = Me.txtTrackingNumber.Value
     .Offset(lastrow, 2).Value = Me.txtOrderNumber.Value
     .Offset(lastrow, 3).Value = Me.lstReturnType.Value
     If Me.chkbxElectronic.Value = True Then
        .Offset(lastrow, 4).Value = Me.txtElecSerial.Value
     End If
     .Offset(lastrow, 5).Value = Me.txtUpc.Value
     .Offset(lastrow, 6).Value = Me.txtSku.Value
     .Offset(lastrow, 7).Value = Me.txtDescription.Value
     .Offset(lastrow, 8).Value = Me.txtLoc.Value
     .Offset(lastrow, 9).Value = Me.txtQty.Value
     If Me.chkbxRestock.Value = False Then
        .Offset(lastrow, 10).Value = "No"
        .Offset(lastrow, 11).Value = Me.lstNoRestock.Value
     Else
        .Offset(lastrow, 10).Value = "Yes"
     End If
     .Offset(lastrow, 12).Value = Me.txtNotes.Value
    End With
     
     
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()

    Dim cCont As Control
    
    For Each cCont In Me.Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.Value = ""
        End If
    Next cCont
    
    Me.lstReturnType.ListIndex = 0
    Me.chkbxRestock.Value = True
    Me.lstNoRestock = Null
    Me.chkbxElectronic.Value = False
    
    

End Sub

Private Sub chkbxElectronic_Change()
    
    lblElecSerial.Visible = chkbxElectronic.Value
    txtElecSerial.Visible = chkbxElectronic.Value

End Sub

Private Sub chkbxRestock_Change()

    lblNoRestock.Visible = Not chkbxRestock.Value
    lstNoRestock.Visible = Not chkbxRestock.Value

End Sub

Private Sub txtUpc_AfterUpdate()

    'Me.txtUpc.Value = Format(Me.txtUpc.Value, "00000000000")
     
    Dim upc As Range
'    Set upc = Workbooks("ReturnsInventory").Sheets("inv").Range("A:A").Find(Me.txtUpc.Value)
    
    With Workbooks("ReturnsInventory").Sheets("inv").Range("A:A")
        Set upc = .Find(Me.txtUpc.Value, LookIn:=xlValues)
    End With
    
    
    MsgBox (upc)

    If upc Then
        Me.txtSku.Value = Application.WorksheetFunction.VLookup(upc, Workbooks("ReturnsInventory").Sheets("inv").Range("A:E"), 2, False)
        Me.txtDescription.Value = Application.WorksheetFunction.VLookup(upc, Workbooks("ReturnsInventory").Sheets("inv").Range("A:E"), 3, False)
        Me.txtLoc.Value = Application.WorksheetFunction.VLookup(upc, Workbooks("ReturnsInventory").Sheets("inv").Range("A:E"), 4, False)
    End If

End Sub

Private Sub UserForm_Initialize()

    With lstReturnType

        lstReturnType.AddItem ("RAVR")
        lstReturnType.AddItem ("RADE")
        lstReturnType.AddItem ("RAIR")
        lstReturnType.AddItem ("Unknown")
        .ListIndex = 0
    
    End With
    
    With lstReturnType

        lstNoRestock.AddItem ("NA Location")
        lstNoRestock.AddItem ("Electronics Return")
        lstNoRestock.AddItem ("Defective Item")
        lstNoRestock.AddItem ("Used Item")
        lstNoRestock.AddItem ("Other")
    
    End With
    
    lblElecSerial.Visible = False
    txtElecSerial.Visible = False
    
    chkbxRestock.Value = True
    
    lblNoRestock.Visible = False
    lstNoRestock.Visible = False
       


End Sub


