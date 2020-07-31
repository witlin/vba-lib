Attribute VB_Name = "Module2"
Sub Summarize()
    ' Declare variables.
    Dim c As Integer             ' repeat part counter.
    Dim i As Integer             ' current cell of the current row in inventory and summary work-sheets.
    Dim sumRow As Integer        ' current row in the summary work-sheet.
    Dim lastPart As String       ' last part number read.
    Dim nwPart As String         ' new part number read (current iteration).
    
    ' Initialize variables.
    c = 1
    lastPart = "null"
    nwPart = "null"
    Set awb = ActiveWorkbook
    Set actSheet = ActiveSheet
    Set sumSheet = awb.WorkSheets("Summary")

    If actSheet.Name = "Inventory" Then
    ' TODO - Find the dynamic beginning of the used range object, instead of the literal A3:...
      Set s = actSheet.Range("A3:T" & sumSheet.UsedRange.Rows.Count - 2)
      Set z = sumSheet.Range("A3:T" & sumSheet.UsedRange.Rows.Count - 2)
      Set cl = sumSheet.Range("A4:T" & sumSheet.UsedRange.Rows.Count - 2)

      'TODO - Add pre-sort and post-sort features in both the inventory and summary worksheets.
      MsgBox "sort by part# in the Inventory work-sheet (column D) before summarizing..." & sumSheet.UsedRange.Rows.Count
      
      ' Clean up Summary sheet from previous use.
      For Each r In cl.Rows
        cl.Rows(r.Row).clearContents
      Next

      ' Loop through all rows.
      For Each rw In s.Rows  
        ' capture the current value in the part column on the inventory sheet
        nwPart = s.Cells(rw.Row, 4).Value

        Debug.Print "inventory row number: " & CStr(rw.Row) & " - new: " & nwPart & " - last: " & lastPart
        
        'if the current part value capture is not equal to the last one run this
        If nwPart <> lastPart Then
          i = 1
          'loop through all the cells in the current row and pass the value from the corresponding cell in the inventory sheet
          'to the matching one in the summary sheet.
          For Each cell In rw.Cells
            z.Cells(rw.Row, i).Value = s.Cells(rw.Row, i).Value
            i = i + 1
          Next
          Debug.Print "new part number: " & s.Cells(rw.Row, 4).Value & " - s qty: " & CStr(s.Cells(rw.Row, 9).Value) & "  - z qty: " & CStr(z.Cells(rw.Row, 9))
          c = 1
        'otherwise run this
        Else 
          'Every time we get a row with a repeated part field we operate in the row of the summary sheet which would be the previous 
          'row in the inventory sheet 
          sumRow = rw.Row - c
          
          'try to aggregate the cell's value with the previous iteration's value, if the input value is a real number then aggregate, otherwise
          'apply a zero value to corresponding summary sheet cell.
          
          'I
          If IsNumeric(s.Cells(rw.Row, 9).Value) Then 
            z.Cells(sumRow, 9).Value = z.Cells(sumRow, 9).Value + s.Cells(rw.Row, 9).Value
          Else
            z.Cells(sumRow, 9).Value = 0
          End If
          'M
          If IsNumeric(s.Cells(rw.Row, 13).Value) Then
            z.Cells(sumRow, 13).Value = z.Cells(sumRow, 13).Value + s.Cells(rw.Row, 13).Value
          Else
            z.Cells(sumRow, 13).Value = 0
          End If
          'O
          If IsNumeric(s.Cells(rw.Row, 15).Value) Then
            z.Cells(sumRow, 15).Value = z.Cells(sumRow, 15).Value + s.Cells(rw.Row, 15).Value
          Else
           z.Cells(sumRow, 15).Value = 0
          End If
          'Q
          If IsNumeric(s.Cells(rw.Row, 17).Value) Then
            z.Cells(sumRow, 17).Value = z.Cells(sumRow, 17).Value + s.Cells(rw.Row, 17).Value   
          Else
           z.Cells(sumRow, 17).Value = 0
          End If
          'S
          If IsNumeric(s.Cells(rw.Row, 19).Value) Then
            z.Cells(sumRow, 19).Value = z.Cells(sumRow, 19).Value + s.Cells(rw.Row, 19).Value   
          Else
           z.Cells(sumRow, 19).Value = 0
          End If
          'U
          If IsNumeric(s.Cells(rw.Row, 21).Value) Then
            z.Cells(sumRow, 21).Value = z.Cells(sumRow, 21).Value + s.Cells(rw.Row, 21).Value   
          Else
           z.Cells(sumRow, 21).Value = 0
          End If
          'now that all values are zero or real numbers, add them up to get total qty received.
          z.Cells(sumRow, 10).Value = z.Cells(sumRow, 13).Value + z.Cells(sumRow, 15).Value + z.Cells(Row, 17).Value + z.Cells(Row, 19).Value + z.Cells(sumRow, 21).Value
          'divide qty received over total qty
          z.Cells(sumRow,11).Value = z.Cells(sumRow, 10) / z.Cells(sumRow, 9)
          
          Debug.Print "    summary row number: " & CStr(sumRow) & " - s qty: " & CStr(s.Cells(rw.Row, 9).Value) & "  - z qty: " & CStr(z.Cells(sumRow, 9).Value)
          
          'we increase the number of repeated parts we already encountered so can continue to operate on the current inventory row and stay in the
          'first row where the current part number wasn't repeated on the summary sheet.
          c = c + 1
        End If
        
        'store the current part number in the last-part number field so we can compare on the next iteration.
        lastPart = nwPart

        'we want to clear the contents of certain fields but not on the header
        If rw.Row > 3 Then
          z.Cells(rw.Row, 1).Value = ""
          z.Cells(rw.Row, 2).Value = ""
          z.Cells(rw.Row, 3).Value = ""
          z.Cells(rw.Row, 8).Value = ""
          z.Cells(rw.Row, 7).Value = ""
        End If
      Next
    Else
      MsgBox "wrong worksheet, go to Inventory" & actSheet.Name
    End If
End Sub
