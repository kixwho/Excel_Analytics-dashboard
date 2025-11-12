Attribute VB_Name = "Module1"
Sub Prepare_Sales_Data()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tbl As ListObject
    Dim dataRange As Range
    
    ' Work on the active sheet
    Set ws = ActiveSheet
    
    ' Find last row in column A (assumes column A always has data)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Format column E as currency
    ws.Columns("E").NumberFormat = "$#,##0.00"
    
    ' Insert new column P for MntTotal
    ws.Columns("P").Insert Shift:=xlToRight
    ws.Range("P1").Value = "MntTotal"
    
    ' Add SUM formula in P2 and fill down dynamically
    ws.Range("P2").FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"
    ws.Range("P2").AutoFill Destination:=ws.Range("P2:P" & lastRow)
    
    ' Define data range dynamically (A1:AC + last row)
    Set dataRange = ws.Range("A1:AC" & lastRow)
    
    ' Remove any existing table (optional)
    On Error Resume Next
    ws.ListObjects(1).Unlist
    On Error GoTo 0
    
    ' Create new table
    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tbl.Name = "SalesData"
    
    ' Sort by MntTotal descending
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("P2:P" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
    
    ' Conditional formatting for column AB > 0.5
    With ws.Columns("AB")
        .FormatConditions.Delete
        With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.5")
            .Font.Color = RGB(0, 102, 0)
            .Interior.Color = RGB(255, 199, 206)
        End With
    End With
    
    ' Insert 3 header rows at top
    ws.Rows("1:3").Insert Shift:=xlDown
    
    ' Add report title and date
    ws.Range("A1").Value = "Monthly Report"
    ws.Range("A2").Value = "Date"
    ws.Range("B2").Formula = "=TODAY()"
    
    MsgBox "Sales data prepared successfully on '" & ws.Name & "'.", vbInformation
End Sub

