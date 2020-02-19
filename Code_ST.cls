VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Performance_MultiYear()
    
    For Each ws In Worksheets
        
        Dim WorksheetName As String
        'grab worksheet name (may not use this)
        WorksheetName = ws.Name
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("We have " + Str(lastrow) + " Rows")
        
        ' Set a variable for holding the Ticker name
        Dim Ticker As String
        
        ' Set a variable for holding the total volume per ticker
        Dim Volume_Total As Double
        Volume_Total = 0
        
        ' Set a variable for holding the Best, Worst performance, Most volume among stocks
        Dim Ticker_Worst As String
        Dim Ticker_Best As String
        Dim Ticker_LargestVolume As String
        
        Dim Percent_Change_Lowest As Double
        Dim Percent_Change_Highest As Double
        Dim Largest_Total_Volume As Double
        Percent_Change_Lowest = 0
        Percent_Change_Highest = 0
        Largest_Total_Volume = 0
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        ' Keep track of the location for each ticker in the Summary table (Columns J-M)
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Set year start row for the first Ticker
        YearStartRow = 2
        
        ' Loop through all entries (Note: separated by ticker using If statement below)
        For i = 2 To lastrow
        
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                YearEndRow = i
                ' Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                
                ' Add to the Volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                Year_Open_Price = ws.Cells(YearStartRow, 3).Value
                Year_Close_Price = ws.Cells(i, 6).Value
                
                Yearly_Change = Year_Close_Price - Year_Open_Price
                ''Noticed a Runtime Error 6 Overflow; This was because Stock year open price was Zero,
                ''leading to division by zero when computing percent change; Used an 'if' below to fix
                If Year_Open_Price <> 0 Then
                    Yearly_Change_Percent = (Year_Close_Price - Year_Open_Price) / Year_Open_Price
                Else
                    Yearly_Change_Percent = 0
                End If
                ws.Range("J" & Summary_Table_Row).Value = Ticker
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("L" & Summary_Table_Row).Value = Yearly_Change_Percent
                ws.Range("M" & Summary_Table_Row).Value = Volume_Total
                
                'Formatting for Summary Row
                ws.Cells(Summary_Table_Row, 12).Style = "Percent"
                'Cells(Summary_Table_Row, 12) = Format(Yearly_Change_Percent, "0.00%")
                
                If ws.Cells(Summary_Table_Row, 11).Value < 0 Then
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                End If
                
                ' Update Best, Worst performing, Largest volume stock list if needed
                If Yearly_Change_Percent < Percent_Change_Lowest Then
                    Percent_Change_Lowest = Yearly_Change_Percent
                    Ticker_Worst = Ticker
                ElseIf Yearly_Change_Percent > Percent_Change_Highest Then
                    Percent_Change_Highest = Yearly_Change_Percent
                    Ticker_Best = Ticker
                End If
                
                If Volume_Total > Largest_Total_Volume Then
                    Largest_Total_Volume = Volume_Total
                    Ticker_LargestVolume = Ticker
                End If
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the Volume Total
                Volume_Total = 0
                
                ' Reset the year start row for the next Ticker
                YearStartRow = i + 1
                
     
            ' If the cell immediately following prev row is the same ticker...
            Else
            
                ' Add to the Volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        'Write Best, Work Performing and Most Volume Stock Info in P1:R4
        
        'ws.Range("P1").Value = ""
        ws.Range("P1").Value = WorksheetName + "-Performance"
        
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Most Volume"
        
        
        ws.Range("Q2").Value = Ticker_Best
        ws.Range("Q3").Value = Ticker_Worst
        ws.Range("Q4").Value = Ticker_LargestVolume
        
        ws.Range("R2").Value = Percent_Change_Highest
        ws.Range("R3").Value = Percent_Change_Lowest
        ws.Range("R4").Value = Largest_Total_Volume
        
        ws.Cells(2, 18).Style = "Percent"
        ws.Cells(3, 18).Style = "Percent"
        
        ws.Range("A1:Z1").Font.Bold = True
        ws.Columns("A:Z").AutoFit
        ws.Columns("A:Z").HorizontalAlignment = xlCenter
    
    Next ws
    
End Sub

