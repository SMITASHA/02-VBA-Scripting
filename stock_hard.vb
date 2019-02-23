Sub Stock_HardCal():

    'Declare Variables
   
    Dim col As Integer
    Dim ws As Worksheet
    Dim ws_count As Integer
    Dim LastRow As Double
    Dim LastRowNewTable As Double
    Dim row_current As Double
    Dim row_current_v As Double
    Dim rng As Range
    Dim rng_v As Range
    Dim cell As Range
    Dim cell_v As Range
    Dim cnt As Integer
    Dim i As Double
    Dim x As Integer
    Dim y As Integer
    Dim LastCol As Integer
    Dim Tot_Volume As Double
    Dim stk_yr_Open As Double
    Dim stk_yr_Close As Double
    Dim yearly_Change As Double
    Dim greatest_percent_Increase As Double
    Dim greatest_percent_Decrease As Double
    Dim greatest_total_volume As Double


    'set worksheet count
    ws_count = ActiveWorkbook.Worksheets.Count
    MsgBox ("Total Number of worksheet: " & Str(ws_count))
    
    'Loop thoroug each Worksheet in Workbook
    For Each ws In Worksheets
        'Activate the current Worksheet
        ws.Activate

        
        'Find the last non-blank cell in column 1
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'Find the last non-blank cell in row 1
        LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        'Initilize variables
      
        col = 1
        cnt = 1
        x = 1
        y = LastCol + 2
        greatest_percent_Increase = 0
        greatest_percent_Decrease = 0
        greatest_total_volume = 0
        
        'initilizing the first row of new table1 in each worksheet
        Cells(x, y).Value = "Ticker"
        Cells(x, y + 1).Value = "Yearly Change"
        'Cells(x, y + 1).WrapText = True
        Cells(x, y + 2).Value = "Percent change"
        'Cells(x, y + 2).WrapText = True
        Cells(x, y + 3).Value = "Total Stock Volume"
        'Cells(x, y + 3).WrapText = True
        x = x + 1
        Total_Vol = 0
        stk_yr_Open = 0
        stk_yr_Close = 0
        
        'initilizing the first row and then column of new table2 in each worksheet
        Cells(1, y + 7).Value = "Ticker"
        Cells(1, y + 8).Value = "Value"
        Cells(2, y + 6).Value = "Greatest % Increase"
        'Cells(2, y + 6).WrapText = True
        Cells(3, y + 6).Value = "Greatest % Decrease"
        'Cells(3, y + 6).WrapText = True
        Cells(4, y + 6).Value = "Greatest total volume"
        'Cells(4, y + 6).WrapText = True
        
        For i = 2 To LastRow
            If cnt = 1 Then
            stk_yr_Open = Cells(i, 3).Value
            End If
            'Year1 = Cells(i, 2).Value
            'year2 = Cells(i + 1, 2).Value
            'if  Left(Year, 4)<>Left(Year, 4) then
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Cells(x, y).Value = Cells(i, 1).Value
                Total_Vol = Total_Vol + Cells(i, 7).Value
                stk_yr_Close = Cells(i, 6).Value
                
            'If stock open price is 0 which is impossible then Yearly change
            'and Percentage change both becomes Not Applicable
                             
                If stk_yr_Open = 0 Then
                Cells(x, y + 1).Value = "N/A"
                Cells(x, y + 2).Value = "N/A"
                'Cells(x, y + 1).Value = stk_yr_Close - stk_yr_Open
                'Cells(x, y + 2).Value = (stk_yr_Close - stk_yr_Open)
                Else
                    yearly_Change = stk_yr_Close - stk_yr_Open
                    Cells(x, y + 1).Value = yearly_Change
                    'Cells(x, y + 1).NumberFormat = "0.000000000"
                    Cells(x, y + 2).Value = yearly_Change / stk_yr_Open
                    ' percent format for percent change
                    Cells(x, y + 2).NumberFormat = "0.00%"
                    
                    If yearly_Change >= 0 Then
                      Cells(x, y + 1).Interior.ColorIndex = 4
                    Else
                    Cells(x, y + 1).Interior.ColorIndex = 3
                    End If
                    
                End If
                            
                Cells(x, y + 3).Value = Total_Vol
                'Cells(x, y + 3).NumberFormat = "#,##0"
                
                Total_Vol = 0
                stk_yr_Open = 0
                stk_yr_Close = 0
                cnt = 1
                x = x + 1
            Else
                Total_Vol = Total_Vol + Cells(i, 7).Value
                cnt = cnt + 1
            End If
        Next i
       
        'Find the last non-blank cell in column y of the new table1
        LastRowNewTable = Cells(Rows.Count, y).End(xlUp).Row
        
        'Set range from which to determine largest and smallest percentage value of the new table1
        Set rng = Range(Cells(2, y + 2), Cells(LastRowNewTable, y + 2))
        Set rng_v = Range(Cells(2, y + 3), Cells(LastRowNewTable, y + 3))
        ' Find the max & min value of percent change from table 1
        greatest_percent_Increase = Application.WorksheetFunction.Max(rng)
        greatest_percent_Decrease = Application.WorksheetFunction.Min(rng)
        For Each cell In rng
            row_current = 0
            row_current = cell.Row
            'find the row which contains the max%change and its ticker value
            If cell.Value = greatest_percent_Increase Then
                'assign the ticker value to the table2 against the max%increase
                Cells(2, y + 7).Value = Cells(row_current, y).Value
            ElseIf cell.Value = greatest_percent_Decrease Then
                'assign the ticker value to the table2 against the min%increase
                Cells(3, y + 7).Value = Cells(row_current, y).Value
            Else
                'Do nothing
            End If
        Next cell
       
        greatest_total_volume = Application.WorksheetFunction.Max(rng_v)
        For Each cell_v In rng_v
            'find the row corrosponding to the max_volume
            If cell_v.Value = greatest_total_volume Then
                row_current_v = cell_v.Row
                'assign the ticker value to the table2 against the max volume
                Cells(4, y + 7).Value = Cells(row_current_v, y).Value
            End If
        Next cell_v
        ' assign the values of table2
        Cells(2, y + 8).Value = greatest_percent_Increase
        Cells(2, y + 8).NumberFormat = "0.00%"
        Cells(3, y + 8).Value = greatest_percent_Decrease
        Cells(3, y + 8).NumberFormat = "0.00%"
        Cells(4, y + 8).Value = greatest_total_volume
        'Cells(x, y + 8).NumberFormat = "#,##0"
        ws.Columns("I:R").AutoFit
        MsgBox ("worksheet No : " & ActiveSheet.Name)
        'MsgBox ("New table : " & Str(greatest_percent_Increase) & Str(greatest_percent_Decrease) & Str(greatest_total_volume))
    Next
End Sub

