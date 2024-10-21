VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FormatDateInColumnF()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim cellValue As Variant
    Dim formattedDate As String
    
    ' Set the worksheet (replace "Sheet1" with your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Loop through each cell in column F
    For Each cell In ws.Range("F2:F" & lastRow)
        cellValue = cell.Value
        
        ' Check if the cell is not blank
        If cellValue <> "" Then
            ' Ensure the cell contains a numeric value and exactly 7 or 8 digits
            If IsNumeric(cellValue) And (Len(cellValue) = 7 Or Len(cellValue) = 8) Then
                On Error GoTo DateError ' Error handling in case DateSerial fails
                ' Format the date in dd/mm/yyyy format
                formattedDate = Format(DateSerial(Mid(cellValue, Len(cellValue) - 3, 4), _
                                                  Mid(cellValue, Len(cellValue) - 5, 2), _
                                                  Mid(cellValue, Len(cellValue) - 7, 2)), "dd/mm/yyyy")
                cell.Value = formattedDate
                On Error GoTo 0 ' Reset error handling
            Else
                ' If not numeric or not the correct length, clear the cell
                cell.Value = ""
            End If
        End If
    Next cell
    Exit Sub
    
DateError:
    MsgBox "Error formatting date in cell " & cell.Address & ": " & Err.Description
    cell.Value = "" ' Clear the cell if there's an error
    Resume Next ' Continue with the next cell
End Sub


