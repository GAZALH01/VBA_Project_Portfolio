Attribute VB_Name = "Module1"
Sub CalculateFIFOValuation()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Inventory() As Variant
   
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Ensure "Sheet1" is the correct name
    
    ' Calculate the last row with data in column A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    
    ' Ensure LastRow is valid
    If LastRow < 2 Then
        MsgBox "No data found or only header present in column A."
        Exit Sub
    End If


    ' Resize the Inventory array
    ReDim Inventory(1 To LastRow - 1, 1 To 3) ' Subtract 1 to account for the header

    
     ' Clear any previous FIFO valuation
    ws.Range("H2:H" & LastRow).ClearContents
    
    ' Loop through each transaction
    For i = 2 To LastRow
        Product = ws.Cells(i, 1).Value
        Quantity = ws.Cells(i, 4).Value
        Cost = ws.Cells(i, 5).Value
    
    
    If ws.Cells(i, 3).Value = "Purchase" Then
        ' Store the purchase in the Inventory array
    For j = 1 To LastRow
        If Inventory(j, 1) = "" Then
            Inventory(j, 1) = Product
            Inventory(j, 2) = Quantity
            Inventory(j, 3) = Cost
            Exit For
        End If
    Next j
ElseIf ws.Cells(i, 3).Value = "Sale" Then
    SaleQty = -Quantity
    FIFOValue = 0
    
    
    ' Apply FIFO logic
    For j = 1 To LastRow
        If Inventory(j, 1) = Product And Inventory(j, 2) > 0 Then
            If Inventory(j, 2) >= SaleQty Then
                FIFOValue = FIFOValue + SaleQty * Inventory(j, 3)
                Inventory(j, 2) = Inventory(j, 2) - SaleQty
                SaleQty = 0
                Exit For
            Else
                FIFOValue = FIFOValue + Inventory(j, 2) * Inventory(j, 3)
                SaleQty = SaleQty - Inventory(j, 2)
                Inventory(j, 2) = 0
            End If
        End If
    Next j
    
    
    ' Record FIFO valuation for the sale
    ws.Cells(i, 8).Value = FIFOValue
    End If
Next i

MsgBox "FIFO Valuation Calculated"

End Sub
