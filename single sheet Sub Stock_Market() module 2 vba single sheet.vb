Sub Stock_Market()

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Dim Ticker as String
        Dim Quarterly_Change as Double
        Quarterly_Change = 0

        Dim Percent_Change as Double
        Percent_Change = 0

        Dim Stock_Volume as Double
        Stock_Volume = 0

        Dim stock_summary as Double
        stock_summary = 2

        Dim Open_price as double
        Dim Close_price as double
    

        LastRow = Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 to LastRow

            If Cells(i+1, 1).Value <> Cells(i, 1).Value Then

                Ticker = Cells(i,1).Value

                Open_price = Cells(i,3).Value

                Close_price = Cells(i,6).Value

                Quarterly_Change = Close_price - Open_price

                Percent_Change = (Quarterly_Change / Open_price) * 100

                Range("K" & stock_summary).Value = Percent_Change 

                Range("J" & stock_summary).Value = Quarterly_Change

                Stock_Volume = Stock_Volume + Cells(i,7).Value

                Range("I" & stock_summary).Value = Ticker

                Range("L" & stock_summary).Value = Stock_Volume

                Range("K" & stock_summary).NumberFormat = "0.00%"

                stock_summary = stock_summary + 1

                Stock_Volume = 0
            Else

                Stock_Volume = Stock_Volume + Cells(i, 7).Value
            End If

            If Quarterly_Change > 0 then
                Range("J" & stock_summary).Interior.ColorIndex = 4
            elseif Quarterly_Change < 0 then
                Range("J" & stock_summary).Interior.ColorIndex = 3
            End if    

        Next I
        
        Dim Greatest_Increase as double
        Dim Greatest_Decrease as double
        Dim Greatest_Total as double

        Greatest_Increase = WorksheetFunction.Max(Range("K:K"))
        Greatest_Decrease = WorksheetFunction.Min(Range("K:K"))
        Greatest_Total = WorksheetFunction.Max(Range("L:L"))

        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O2").Value = Greatest_Increase
        Range("O3").Value = Greatest_Decrease
        Range("O4").Value = Greatest_Total
        Range("O2").NumberFormat = "0.00%"
        Range("O3").NumberFormat = "0.00%"
       
End Sub            