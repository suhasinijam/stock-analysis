Sub module2Assignment()
    ' Declaring worksheet and variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim total_stock_volume As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim ticker_location As Long
    Dim start As Double
    Dim lastrow As Long
    Dim i As Long
   
   
   
    ' Worksheet for-loop
    For Each ws In ThisWorkbook.Worksheets
   
    'declaring variables
    ws.Cells(1, 9).value = "ticker"
    ws.Cells(1, 12).value = "total_stock_volume"
    ws.Cells(1, 10).value = "yearly_change"
    ws.Cells(1, 11).value = "percentage_change"
   
    ' Fixing the value so that this can be used to allocate the next ticker in column I
    ticker_location = 2
       
     ' Formula to declare last row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Reset variables for each worksheet
        total_stock_volume = 0
        start = Cells(2, 3).value
       
        ' Start of for-loop to enter values
        For i = 2 To lastrow
            If ws.Cells(i, 1).value <> ws.Cells(i + 1, 1).value Then
                ' Ticker name
                ticker = ws.Cells(i, 1).value
               
                ' Allocating ticker name
                ws.Cells(ticker_location, 9).value = ticker
               
                'yearly change
                yearly_change = Cells(i, 6).value - start
                'allocating to yearly change column
                ws.Cells(ticker_location, 10).value = yearly_change
               
                'calculating percentage change
                ws.Cells(ticker_location, 11).value = yearly_change / start
               
                ' Resetting start value to new open value
                start = ws.Cells(i + 1, 3).value
               
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).value
               
                'allocating total_stock_volue
               ws.Cells(ticker_location, 12).value = total_stock_volume
                ' Resetting total stock volume
                total_stock_volume = 0
               
                ' Increment the ticker_location for the next record
                ticker_location = ticker_location + 1
               
            Else
                ' Calculating total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).value
               
                'allocating total_stock_volue
                'total_stock_volume = ws.Cells(ticker_location, 12).Value
           
            End If
        Next i

    Next ws
End Sub
