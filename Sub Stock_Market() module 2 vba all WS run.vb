Sub Stock_Market()
    For Each ws in Worksheets

        Dim WorksheetName as String

        WorksheetName = ws.Name

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim Ticker as String
        Dim Quarterly_Change as Double
        Quarterly_Change = 0

        Dim Percent_Change as Double
        Percent_Change = 0

        Dim Stock_Volume as Double
        Stock_Volume = 0

        Dim stock_summary as Double
        stock_summary = 2

        Dim Open_price as Double
        Dim Close_price as Double
    

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 to LastRow

            If ws.Cells(i+1, 1).Value <> ws.Cells(i, 1).Value Then

                Ticker = ws.Cells(i,1).Value

                Open_price = ws.Cells(i,3).Value

                Close_price = ws.Cells(i,6).Value

                Quarterly_Change = (Close_price - Open_price)

                Percent_Change = ((Quarterly_Change / Open_price)*100)

                ws.Range("K" & stock_summary).Value = Percent_Change 

                ws.Range("J" & stock_summary).Value = Quarterly_Change

                Stock_Volume = Stock_Volume + ws.Cells(i,7).Value

                ws.Range("I" & stock_summary).Value = Ticker

                ws.Range("L" & stock_summary).Value = Stock_Volume

                ws.Range("K" & stock_summary).NumberFormat = "0.00%"

                stock_summary = stock_summary + 1

                Stock_Volume = 0
            Else

                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            End If
            
            If Quarterly_Change > 0 then
                ws.Range("J" & stock_summary).Interior.ColorIndex = 4
            elseif Quarterly_Change < 0 then
                ws.Range("J" & stock_summary).Interior.ColorIndex = 3
            End if  

        Next I
        
        Dim Greatest_Increase as Double
        Dim Greatest_Decrease as Double
        Dim Greatest_Total as Double

        Greatest_Increase = WorksheetFunction.Max(Range("K:K"))
        Greatest_Decrease = WorksheetFunction.Min(Range("K:K"))
        Greatest_Total = WorksheetFunction.Max(Range("L:L"))

        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O2").Value = Greatest_Increase
        ws.Range("O3").Value = Greatest_Decrease
        ws.Range("O4").Value = Greatest_Total
        ws.Range("O2").NumberFormat = "0.00%"
        ws.Range("O3").NumberFormat = "0.00%"
       
    Next ws    
End Sub            