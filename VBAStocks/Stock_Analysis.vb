Sub Stock_analysis()
    'Set the initial variables for all worksheets
    Dim Ticker As String
    Dim LastRow As Long
    Dim Price_Open As Double
    Dim Price_Close As Double
    Dim Price_Change As Double
    Dim Price_PerctChg As Double
    Dim Stock_Vol As Double
    Dim Output_Row As Integer
    Dim Output_Col() As String
    Dim Column_Names() As String
    Dim Output_Array() As Variant
    Dim Great_Perct_Inc As Variant
    Dim Great_Perct_Dec As Variant
    Dim Great_Total_Vol As Variant
    Dim Final_Output_Header() As String

      
    'Declare ws as a worksheet object variable to be able apply script to all worksheets and be able to write to each worksheet
    Dim ws As Worksheet
     
    
    'Initialize the Arrays to the capture Greatest values from all worksheets
    Great_Perct_Inc = Array("Greatest % Increase", Ticker, 0)
    Great_Perct_Dec = Array("Greatest % Decrease", Ticker, 0)
    Great_Total_Vol = Array("Greatest Total Volume", Ticker, 0)
        
        
    'Loop through all sheets in the workbook
    For Each ws In Worksheets
    
        'MsgBox ws.Name
        
        'Initialize Variables at the start of analyzing each worksheet
        Output_Col = Split("K,L,M,N", ",")
        Column_Names = Split("Ticker Symbol,Yearly Change $,Percent Change %,Total Stock Volume", ",")
        Output_Row = 1
        Stock_Vol = 0
        Price_Open = ws.Cells(2, 3).Value
        Ticker = ws.Cells(2, 1).Value
        Final_Output_Header = Split("Ticker,Value", ",")
        
        'Initialize the Arrays to the capture Greatest values from all worksheets
        Great_Perct_Inc = Array("Greatest % Increase", Ticker, 0)
        Great_Perct_Dec = Array("Greatest % Decrease", Ticker, 0)
        Great_Total_Vol = Array("Greatest Total Volume", Ticker, 0)
        
        'Find last row for current worksheet to determine how many iterations of the For loop
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Write the Output_Col Header values using Column_Names array in a loop
        For i = 0 To UBound(Output_Col)
            ws.Range(Output_Col(i) & Output_Row).Value = Column_Names(i)
        Next i
        
        'Increment Output_Row
        Output_Row = Output_Row + 1
        
        
        'Loop through all the stock entries in the current worksheet starting at row 2 and ending at last row with data (from LastRow variable above)
        For i = 2 To LastRow
        
            'Check if we are on last row of Ticker symbol - if yes then summarize data for that Ticker and write output to current worksheet
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set Ticker Symbol
                Ticker = ws.Cells(i, 1).Value
                
                'Increment Stock_Vol variable
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
                
                'Set Price_Close Variable
                Price_Close = ws.Cells(i, 6).Value
                
                'Calculate Yearly change from opening price at beginning of the year to closing price at end of that year
                Price_Change = Price_Close - Price_Open
                
                'Calculate the Percent Change of the Open vs. Close Stock Price
                'Avoid division by zero error by setting Price_Open to 1 cent in situation that Price_Open = 0
                If Price_Open <> 0 Then
                    Price_PerctChg = Price_Change / Price_Open
                Else
                    Price_PerctChg = Price_Change / 0.01
                End If
                
                'Compare and capture the Greatest Percentage Change in price - add ticker and Price_PerctChg to collection if max or min
                If Price_PerctChg > Great_Perct_Inc(2) Then
                    Great_Perct_Inc(1) = Ticker
                    Great_Perct_Inc(2) = Price_PerctChg
                ElseIf Price_PerctChg < Great_Perct_Dec(2) Then
                    Great_Perct_Dec(1) = Ticker
                    Great_Perct_Dec(2) = Price_PerctChg
                End If
                
                
                'Compare and capture the Greatest Total Volume - add ticker and Stock_Vol if greater than the current Great_Total_Vol
                If Stock_Vol > Great_Total_Vol(2) Then
                    Great_Total_Vol(1) = Ticker
                    Great_Total_Vol(2) = Stock_Vol
                End If
                
                
                
                Output_Array = Array(Ticker, Price_Change, Price_PerctChg, Stock_Vol)
                
                'Write the results to a new Output_Row
                For k = 0 To UBound(Output_Col)
                    If k = 1 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        If Output_Array(k) > 0 Then
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 4
                        ElseIf Output_Array(k) < 0 Then
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 3
                        Else
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 6
                        End If
                    ElseIf k = 2 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        ws.Range(Output_Col(k) & Output_Row).NumberFormat = "#,##0.00%"
                    ElseIf k = 3 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        ws.Range(Output_Col(k) & Output_Row).NumberFormat = "#,###"
                    Else
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                    End If
                Next k
            
                'Reset Variables for next Ticker Symbol
                Ticker = ws.Cells(i + 1, 1).Value
                Stock_Vol = 0
                Price_Open = ws.Cells(i + 1, 3).Value
                Output_Row = Output_Row + 1
                            
                
            'If we are not on last ticker row then increment and set values
            Else
                
                'Add to Total Stock_Vol
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        'Write the final results for the Greatest Values for % Increase, % Decrease and Highest Volume
        ws.Range("R1:S1").Value = Final_Output_Header
        ws.Range("Q2:S2").Value = Great_Perct_Inc
        ws.Range("S2").NumberFormat = "#,##0.00%"
        ws.Range("Q3:S3").Value = Great_Perct_Dec
        ws.Range("S3").NumberFormat = "#,##0.00%"
        ws.Range("Q4:S4").Value = Great_Total_Vol
        ws.Range("S4").NumberFormat = "#,###"
        
        'Fit all the columns to correct size to display widest contents
        ws.Columns("A:Q").EntireColumn.AutoFit
    Next

    
    
End Sub

