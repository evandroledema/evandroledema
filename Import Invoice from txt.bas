Sub importInvoicesRelative()

' Written by Evandro Ledema. Import invoice.txt and add totals.
'April, 17. 2021.

    Dim FinalRow, TotalRow As Integer
    'Declaring all the variable used.
    
    ChDir "C:\Users"
    'The std directory.
    
    Workbooks.OpenText Filename:="C:\Users\invoice2.txt", _
        Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
        , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 3), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)), _
        TrailingMinusNumbers:=True
    'Importing the data from a .txt: Delimited by comma.
    
    FinalRow = Cells(1, 1).CurrentRegion.Rows.Count
    'One way to count all the rows with data.
    
    TotalRow = FinalRow + 1
    'Total row receive the counted rows with data + 1, it will be necessary to enter all
    '   Totals in the last row.
    
    Cells(TotalRow, 1).Value = "Total"
    'Inserts total label in the next row after the data
    
    Cells(TotalRow, 5).Resize(1, 3).FormulaR1C1 = "=SUM(R2C:R[-1]C)"
    'Inserts the sum formula in the columns 5 to 7 using resize
    
    Cells(TotalRow, 1).Resize(1, 7).Font.Bold = True
    'Makes the cells bold in the lst row, from column 1 through 7
    
    With Cells(1, 1).Resize(1, 7).Font
    .Bold = True
    .Size = 16
    .ColorIndex = 5
    .Underline = xlUnderlineStyleDoubleAccounting
    End With
    'Formats the first row using with and resize
    
    Cells(1, 1).Resize(TotalRow, 7).Columns.AutoFit
    'Autofit the data using resize
    
    Calculate
    'It is used to calculate the formulas in the worksheets with manual formulas
    
    
End Sub
