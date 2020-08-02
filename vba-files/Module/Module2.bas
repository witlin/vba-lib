Attribute VB_Name = "Module2"

Sub Sorter()
  Dim vendorTable, sumTable As ListObject
  Dim vendorRows As ListRows
  Dim dataRange, sumDataRange, range, singlePo, singlePart As Range

  Set vendorTable = ActiveSheet.ListObjects("VendorInventory")
  Set dataRange = vendorTable.DataBodyRange
  Set sumTable = ActiveWorkbook.Worksheets("Summary") _ 
                               .ListObjects("InventorySummary")
  Set sumDataRange = sumTable.DataBodyRange
  Set singlePo = dataRange.Cells(1,1)
  Set singlePart = dataRange.Cells(1,4)

  Debug.Print "inventory data range: " & dataRange.Address
  Debug.Print "summary data range: " & sumDataRange.Address

  For Each obj In ActiveSheet.ListObjects
    Debug.Print obj.Name 
  Next
  For Each obj In ActiveWorkbook.Worksheets("Summary").ListObjects
    Debug.Print obj.Name 
  Next

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
  
  Dim repCounter, cellIter, poBegin, poEnd, sumRow As Integer
  Dim lastPart, newPart, lastPO, thisPO As String
  Dim vendorTable, sumTable As ListObject
  Dim vendorRows, sumRows As ListRows
  Dim dataRange, range, singlePo, singlePart As Range

  repCounter = 1
  lastPart = "null"
  nwPart = "null"

  Set vendorTable = ActiveSheet.ListObjects("VendorInventory")
  Set dataRange = vendorTable.DataBodyRange
  Set sumTable = ActiveWorkbook.Worksheets("Summary") _ 
                               .ListObjects("VendorInventory")
  Set sumDataRange = sumTable.DataBodyRange
  
  'Clear contents in Summary worksheet
  For Each r In dataRange.Rows(r.Row)
    dataRange.Rows(r.Row)
  Next

  For Each r In dataRange.Rows
    newPart = dataRange.Cells(r.Row, 4).Value
    Debug.Print "inventory row number: " & CStr(r.Row) & _
                " - new: " & nwPart & " - last: " & lastPart
    If newPart <> lastPart Then
      cellIter = 1
      For Each cell In r.Cells
            sumDataRange.Cells(rw.Row, cellIter).Value = _ 
              dataRange.Cells(rw.Row, cellIter).Value
            cellIter = cellIter + 1
      Next
      repCounter = 1
      Debug.Print "po: " & dataRange.Cells(r.Row, 1) & _ 
                  "new part: " & dataRange.Cells(r.Row, 4) & _ 
                  "last part: " & dataRange.Cells(r.Row, 9)            
    Else
      sumRow = r.Row - repCounter

    End If
  Next

End Sub