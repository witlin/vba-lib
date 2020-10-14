Attribute VB_Name = "Module2"

Sub SortInventory()
  
  Dim vendorTable, sumTable As ListObject
  Dim vendorRows As ListRows
  Dim invRange, sumRange, range, singlePo, singlePart As Range

  Set vendorTable = ActiveSheet.ListObjects("VendorInventory")
  Set invRange = vendorTable.DataBodyRange
  Set sumTable = ActiveWorkbook.Worksheets("Summary") _ 
                               .ListObjects("InventorySummary")
  Set sumRange = sumTable.DataBodyRange
  Set singlePo = invRange.Cells(1,1)
  Set singlePart = invRange.Cells(1,4)

  Debug.Print "inventory data range: " & invRange.Address
  Debug.Print "summary data range: " & sumRange.Address

  For Each obj In ActiveSheet.ListObjects
    Debug.Print obj.Name 
  Next
  For Each obj In ActiveWorkbook.Worksheets("Summary").ListObjects
    Debug.Print obj.Name 
  Next
  
  ' invRange.Select

  ' sort the inventory table by po first and then by part number.
  vendorTable.Sort. _
      SortFields.Add Key:=singlePo, SortOn:= _
      xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With vendorTable.Sort
      .SortFields.Add Key:=singlePo, SortOn:=xlSortOnValues, _
                      Order:=xlAscending, DataOption:=xlSortNormal
      .SortFields.Add Key:=singlePart, SortOn:=xlSortOnValues, _
                      Order:=xlAscending, DataOption:=xlSortNormal
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
End Sub

Sub Summarize()
  On Error Resume Next

  Dim repCounter, cellIter, sumRow, rowNum As Integer
  Dim lastPart, newPart As String
  Dim vendorTable, sumTable As ListObject
  Dim vendorRows, sumRows As ListRows
  Dim invRange, sumRange, singlePo, singlePart As Range

  repCounter = 1
  lastPart = "null"
  newPart = "null"
  
  ' references to the table objects and the range of their data bodies
  Set vendorTable = ActiveSheet.ListObjects("VendorInventory")
  Set invRange = vendorTable.DataBodyRange
  Set sumTable = ActiveWorkbook.Worksheets("Summary") _ 
                               .ListObjects("InventorySummary")
  Set sumRange = sumTable.DataBodyRange
  ' sumRange.Select

  ' references to the po and part numbers
  Set singlePo = invRange.Cells(1,1)
  Set singlePart = invRange.Cells(1,4)
  
  ' ' clear column filters
  ' For i = 1 to vendorTable.ListColumns.Count 
  '   invRange.Rows.Item(6).Range.AutoFilter Field:=i
  '   sumRange.Rows.Item(6).Range.AutoFilter Field:=i
  ' Next

  ' clear contents in Summary worksheet
  For Each r In sumRange.Rows
    sumRange.Rows(r.Row).ClearContents
  Next

  ' iterate over all rows in the inventory worksheet
  For Each r In invRange.Rows
    rowNum = r.Row - 5
    newPart = invRange.Cells(rowNum, 4).Value
    Debug.Print "inventory row number: " & CStr(rowNum) & _
                " - po: " & invRange.Cells(rowNum, 1) & _ 
                " - new: " & newPart & " - last: " & lastPart
    If newPart <> lastPart Then
      ' if the current row's part number is the same as the part 
      ' number from the last iteration's row, loop through all the 
      ' fields on the current row and assign the values from the row 
      ' in the inventory sheet to the fields of the corresponding row
      ' in the summary sheet.
      cellIter = 1
      For Each cell In r.Cells
            sumRange.Cells(rowNum, cellIter).Value = _ 
              invRange.Cells(rowNum, cellIter).Value
            cellIter = cellIter + 1
      Next
      ' reset the repetition counter to one since the current row hasn't repeated yet
      repCounter = 1          
    Else
      ' sumRow is the index used to append summarized rows to the summary table
      ' it substracts the repetition counter value every iteration of a repeated
      ' row in order to keep adding to the same sumary row.
      sumRow = rowNum - repCounter

          ' try to aggregate the cell's value with the previous iteration's value, 
          ' if the input value is a real number then aggregate, otherwise
          ' apply a zero value to corresponding summary sheet cell.    
          'I
          If IsNumeric(invRange.Cells(rowNum, 9).Value) Then 
            sumRange.Cells(sumRow, 9).Value = _
                  sumRange.Cells(sumRow, 9).Value + _
                  invRange.Cells(rowNum, 9).Value
          Else
            sumRange.Cells(sumRow, 9).Value = 0
          End If
          'M
          If IsNumeric(invRange.Cells(rowNum, 13).Value) Then
            sumRange.Cells(sumRow, 13).Value = _
                  sumRange.Cells(sumRow, 13).Value + _
                  invRange.Cells(rowNum, 13).Value
          Else
            sumRange.Cells(sumRow, 13).Value = 0
          End If
          'O
          If IsNumeric(invRange.Cells(rowNum, 15).Value) Then
            sumRange.Cells(sumRow, 15).Value = _
                  sumRange.Cells(sumRow, 15).Value + _
                  invRange.Cells(rowNum, 15).Value
          Else
           sumRange.Cells(sumRow, 15).Value = 0
          End If
          'Q
          If IsNumeric(invRange.Cells(rowNum, 17).Value) Then
            sumRange.Cells(sumRow, 17).Value = _
                  sumRange.Cells(sumRow, 17).Value + _
                  invRange.Cells(rowNum, 17).Value   
          Else
           sumRange.Cells(sumRow, 17).Value = 0
          End If
          'S
          If IsNumeric(invRange.Cells(rowNum, 19).Value) Then
            sumRange.Cells(sumRow, 19).Value = _
                  sumRange.Cells(sumRow, 19).Value + _
                  invRange.Cells(rowNum, 19).Value   
          Else
           sumRange.Cells(sumRow, 19).Value = 0
          End If
          'U
          If IsNumeric(invRange.Cells(rowNum, 21).Value) Then
            sumRange.Cells(sumRow, 21).Value = _
                  sumRange.Cells(sumRow, 21).Value + _
                  invRange.Cells(rowNum, 21).Value   
          Else
           sumRange.Cells(sumRow, 21).Value = 0
          End If
          ' now that all values are zero or real numbers
          ' add them up to get total qty received
          sumRange.Cells(sumRow, 10).Value = _
                sumRange.Cells(sumRow, 13).Value + _ 
                sumRange.Cells(sumRow, 15).Value + _
                sumRange.Cells(sumRow, 17).Value + _
                sumRange.Cells(sumRow, 19).Value + _
                sumRange.Cells(sumRow, 21).Value
          ' divide qty received over total qty
          sumRange.Cells(sumRow,11).Value = _
                sumRange.Cells(sumRow, 10).Value / _
                sumRange.Cells(sumRow, 9).Value
          
          ' we increase the number of repeated parts we already encountered so 
          ' can continue to operate on the current inventory row and stay in the
          ' first row where the current part number wasn't repeated on the summary sheet
          repCounter = repCounter + 1
    End If        
    'store the current part number in the last-part number field so we can compare on the next iteration
    lastPart = newPart

    'we want to clear the contents of unused columns. TODO - remove unused columns.
    'sumRange.Cells(rowNum, 1).Value = ""
    sumRange.Cells(rowNum, 2).Value = ""
    sumRange.Cells(rowNum, 3).Value = ""
    sumRange.Cells(rowNum, 8).Value = ""
    sumRange.Cells(rowNum, 7).Value = ""
  Next

  ' sort the summary table by po 
  vendorTable.Sort. _
      SortFields.Add Key:=singlePo, SortOn:= _
      xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With vendorTable.Sort
      .SortFields.Add Key:=singlePart, SortOn:=xlSortOnValues, _
                      Order:=xlAscending, DataOption:=xlSortNormal
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
End Sub