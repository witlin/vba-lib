Attribute VB_Name = "Module2"
Sub Summarize()
    ' Declare variables
    Dim c As Integer             ' repeat part counter.
    Dim i As Integer             ' current cell of the current row in inventory and summary work-sheets.
    Dim sumRow As Integer        ' current row in the summary work-sheet.
    Dim lastPart As String       ' last part number read.
    Dim nwPart As String         ' new part number read (current iteration).
    
    ' Initialize variables
    c = 1
    lastPart = "null"
    nwPart = "null"
    Set awb = ActiveWorkbook
    Set actSheet = ActiveSheet
    Set sumSheet = awb.WorkSheets("Summary")

    If actSheet.Name = "Inventory" Then
      Set s = actSheet.Range("A4:T" & sumSheet.UsedRange.Rows.Count - 1)
      Set z = sumSheet.Range("A4:T" & sumSheet.UsedRange.Rows.Count - 1)

      'TODO - Add auto-sort feature
      MsgBox "sort by part# in the Inventory work-sheet (column D) before summarizing..." & sumSheet.UsedRange.Rows.Count
      
      ' Loop through all rows
      For Each rw In s.Rows  
        nwPart = s.Cells(rw.Row, 4).Value
        Debug.Print "inventory row number: " & CStr(rw.Row) & " - new: " & nwPart & " - last: " & lastPart
        s.Cells(rw.Row, 8).Value = ""
        s.Cells(rw.Row, 7).Value = ""
        If nwPart <> lastPart Then
          i = 1
          For Each cell In rw.Cells
            z.Cells(rw.Row, i).Value = s.Cells(rw.Row, i).Value
            i = i + 1
          Next
          Debug.Print "new part number: " & s.Cells(rw.Row, 4).Value & " - s qty: " & CStr(s.Cells(rw.Row, 9).Value) & "  - z qty: " & CStr(z.Cells(sumRow, 9))
          c = 1
        Else
          sumRow = rw.Row - c
          z.Cells(sumRow, 9).Value = z.Cells(sumRow, 9).Value + s.Cells(rw.Row, 9).Value
          Debug.Print "    summary row number: " & CStr(sumRow) & " - s qty: " & CStr(s.Cells(rw.Row, 9).Value) & "  - z qty: " & CStr(z.Cells(sumRow, 9).Value)
          c = c + 1
        End If
        lastPart = nwPart
      Next
    Else
      MsgBox "wrong worksheet, go to Inventory" & actSheet.Name
    End If
End Sub
