VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_year_stock_data()
        
    'macro works across all worksheets
    For Each ws In Worksheets
    
        'declare variables. some variables are declared as double instead of long to prevent overflow errors
        Dim lr As Double
        Dim vol_count As Double
        Dim row_count As Long
        Dim opening As Double
        Dim closing As Double
        
        'set values for last row, volume count, and row counter
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
        vol_count = 0
        row_count = 2
        
        '1st for loop from to generate summary table
        For i = 2 To lr
            
            'set opening and closing variables by searching ticker column for first and last appearances of each unique ticker
            opening = ws.Range("A1:A" & lr).Find(What:=ws.Cells(i, 1).Value, lookat:=xlWhole).Row
            closing = ws.Range("A1:A" & lr).Find(What:=ws.Cells(i, 1).Value, lookat:=xlWhole, searchdirection:=xlPrevious).Row
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'if ticker<>ticker below it, print ticker name
                ws.Range("I" & row_count).Value = ws.Cells(i, 1).Value
                
                'calculate yearly change by subtracting closing price from opening price (using "opening" and "closing" variables)
                ws.Range("J" & row_count).Value = ws.Cells(closing, 6).Value - ws.Cells(opening, 3).Value
    
                'calculate percentage change
                ws.Range("K" & row_count).Value = (ws.Range("J" & row_count).Value / ws.Cells(opening, 3).Value)
                ws.Range("K" & row_count).NumberFormat = "0.00%"
                        
                'add volume of last ticker to volume count, and print final volume count
                vol_count = vol_count + ws.Cells(i, 7).Value
                ws.Range("L" & row_count).Value = vol_count
                
                'apply conditional formatting
                If ws.Range("J" & row_count).Value > 0 Then
                    ws.Range("J" & row_count).Interior.ColorIndex = 4
                    ws.Range("K" & row_count).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & row_count).Interior.ColorIndex = 3
                    ws.Range("K" & row_count).Interior.ColorIndex = 3
                End If
                
                'reset volume counter and increase row counter for next loop
                vol_count = 0
                row_count = row_count + 1
            
            Else
                'if ticker = ticker below it, continue adding the volume to the volume counter
                vol_count = vol_count + ws.Cells(i, 7).Value
    
            End If
        
        Next i
        
        'new "last row" variable is declared and set, as the summary table is a different size
        Dim lr2 As Double
        Dim max_pc_inc As Double
        Dim max_pc_dec As Double
        Dim max_vol As Double
    
        lr2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        max_pc_inc = 0
        max_pc_dec = 0
        max_vol = 0
        
        '2nd for loop to calculate 3rd table
        For i = 2 To lr2
            If ws.Cells(i, 11).Value > max_pc_inc Then
                max_pc_inc = ws.Cells(i, 11).Value
                
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = max_pc_inc
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(i, 11).Value < max_pc_dec Then
                max_pc_dec = ws.Cells(i, 11).Value
                
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = max_pc_dec
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(i, 12).Value > max_vol Then
                max_vol = ws.Cells(i, 12).Value
                
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = max_vol
            End If
        Next i
        
        'add column/row headings for both tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'format column widths
        ws.Columns("I").ColumnWidth = 15
        ws.Columns("J").ColumnWidth = 15
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("L").Columns.AutoFit
        ws.Columns("O").Columns.AutoFit
        ws.Columns("P").ColumnWidth = 15
        ws.Columns("Q").Columns.AutoFit
        
    'move onto next worksheet
    Next ws
End Sub
