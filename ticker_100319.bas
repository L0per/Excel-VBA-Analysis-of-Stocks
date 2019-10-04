Attribute VB_Name = "Module11"
Sub Ticker()

'Speed up the code
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'Create headers for tables and assign column numbers to them
    'Get total number of columns for the first sheet
    Dim lastcolumn As Integer
    lastcolumn = Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Create headers as an array
    Dim SumHeaders
    SumHeaders = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume", "Value")
    
    'Assign varibales and summary tables column numbers
    Dim headerarraycount As Integer
    Dim Tickercolumn As Integer
    Dim Yearchangecolumn As Integer
    Dim Percentchangecolumn As Integer
    Dim tickerlimitscolumn As Integer
    Dim valuelimitscolumn As Integer
    Dim greatestrowscolumn As Integer
    Tickercolumn = lastcolumn + 2 'Ticker names
    Yearchangecolumn = lastcolumn + 3 'Change of stock value over the year
    Percentchangecolumn = lastcolumn + 4 'Percent change
    totalcolumn = lastcolumn + 5 'Total volume of stock
    greatestrowscolumn = lastcolumn + 8 'Row headers for limits table
    tickerlimitscolumn = lastcolumn + 9 'Ticker names for limits table
    valuelimitscolumn = lastcolumn + 10 'Values for limits table
    
    
'Copy Data to summary tables
'Worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
    
    'Apply headers to the assigned column numbers
    ws.Cells(1, Tickercolumn).Value = SumHeaders(0)
    ws.Cells(1, Yearchangecolumn).Value = SumHeaders(1)
    ws.Cells(1, Percentchangecolumn).Value = SumHeaders(2)
    ws.Cells(1, totalcolumn).Value = SumHeaders(3)
    ws.Cells(2, greatestrowscolumn).Value = SumHeaders(4)
    ws.Cells(3, greatestrowscolumn).Value = SumHeaders(5)
    ws.Cells(4, greatestrowscolumn).Value = SumHeaders(6)
    ws.Cells(1, tickerlimitscolumn).Value = SumHeaders(0)
    ws.Cells(1, valuelimitscolumn).Value = SumHeaders(7)
    
    'Format percent change cells to percentage
    ws.Columns(Percentchangecolumn).NumberFormat = "0.00%"
    ws.Cells(2, valuelimitscolumn).NumberFormat = "0.00%" 'Greatest increase in limits table
    ws.Cells(3, valuelimitscolumn).NumberFormat = "0.00%" 'Greatest decrease in limits table
    
    'Format change cells to show small decimals, looks weird this way, find how to do floating point
    'ws.Columns(Yearchangecolumn).NumberFormat = "0.0000000000000"
    
    'Create summary row variable that will move down one in the summary table for each new ticker symbol
    Dim summaryrow As Integer
    summaryrow = 2
    
    'Assign total number of rows to a variable
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the range of the first ticker
    Dim currentticker As String
    Dim tickercount As Integer
    currentticker = ws.Cells(2, 1).Value 'Get the ticker symbol
    tickercount = Application.WorksheetFunction.CountIf(Range(ws.Cells(2, 1), ws.Cells(lastrow, 1)), currentticker) 'Get the total number of rows containing the ticker symbol

    'Find the least and greatest value in the date column
    'Note that 2 is added to the ticker count for the end range to compensate for the header
    Dim mintickerdate As Double
    Dim maxtickerdate As Double
    mintickerdate = Application.WorksheetFunction.Min(Range(ws.Cells(2, 2), ws.Cells(2 + tickercount, 2))) 'Finds smallest date value in the range of cells containing the ticker
    maxtickerdate = Application.WorksheetFunction.Max(Range(ws.Cells(2, 2), ws.Cells(2 + tickercount, 2))) 'Same was the above line, but finds the largest
    
    'Find the rows with the min and max dates relative to the range used
    '2 is also added to the ticker count, see above
    'Match function is used to find the exact cell match for the ticker dates
    Dim mintickerdaterow As Double
    Dim maxtickerdaterow As Double
    mintickerdaterow = 1 + Application.WorksheetFunction.Match(mintickerdate, Range(ws.Cells(2, 2), ws.Cells(2 + tickercount, 2)), 0)
    maxtickerdaterow = 1 + Application.WorksheetFunction.Match(maxtickerdate, Range(ws.Cells(2, 2), ws.Cells(2 + tickercount, 2)), 0)
    
    'Enter tickername in summary table
    ws.Cells(summaryrow, Tickercolumn).Value = currentticker
    
    'Calculate the yearly stock change and enter into to summary table
    'Formating is applied whether the change is positive or negative, borders are also colored for readability
    Dim yearchange As Double
    yearchange = ws.Cells(maxtickerdaterow, 6) - ws.Cells(mintickerdaterow, 3)
    ws.Cells(summaryrow, Yearchangecolumn).Value = yearchange
    If yearchange > 0 Then
        ws.Cells(summaryrow, Yearchangecolumn).Interior.ColorIndex = 4
    End If
    If yearchange < 0 Then
        ws.Cells(summaryrow, Yearchangecolumn).Interior.ColorIndex = 3
    End If
    ws.Cells(summaryrow, Yearchangecolumn).Borders.Color = RGB(216, 216, 216)
    
    'Calculate the stock change as a percentage
    Dim percentchange As Double
    percentchange = (ws.Cells(maxtickerdaterow, 6) - ws.Cells(mintickerdaterow, 3)) / ws.Cells(mintickerdaterow, 3)
    ws.Cells(summaryrow, Percentchangecolumn).Value = percentchange
    
    'Sum the total stock volume
    Dim total As Double
    total = Application.WorksheetFunction.Sum(Range(ws.Cells(2, 7), ws.Cells(2 + tickercount, 7)))
    ws.Cells(summaryrow, totalcolumn).Value = total
    
    'Create row variable for tracking the row position of the last ticker entry
    'Not really used, but may be useful
    Dim lasttickerrow As Integer
    lasttickerrow = 1 + tickercount
    
    'Do loop that moves on to next worksheet once the last ticker is entered
    'This is really the first row for each ticker, but I named it tickerstopper since it's used as a trigger to move to the next worksheet
    'While I could write the code to contain a single loop, I found it much easier to get values using the first ticker of each sheet first, then loop over the remaining tickers
    'A majority of the code in the loop is mostly a copy of what was performed for the first ticker
    Dim tickerstopper As Double
    tickerstopper = lasttickerrow + 1
    Do While tickerstopper < lastrow
        
        'Increase summaryrow variable so that the next ticker goes to the next row
        summaryrow = summaryrow + 1
        
        'Find the range of the ticker
        currentticker = ws.Cells(tickerstopper, 1).Value
        tickercount = Application.WorksheetFunction.CountIf(Range(ws.Cells(tickerstopper, 1), ws.Cells(lastrow, 1)), currentticker)
        
    
        'Find the least and greatest value in the date column
        mintickerdate = Application.WorksheetFunction.Min(Range(ws.Cells(tickerstopper, 2), ws.Cells(tickerstopper - 1 + tickercount, 2)))
        maxtickerdate = Application.WorksheetFunction.Max(Range(ws.Cells(tickerstopper, 2), ws.Cells(tickerstopper - 1 + tickercount, 2)))
        
        'Find the rows with the min and max dates relative to the range used
        mintickerdaterow = tickerstopper - 1 + Application.WorksheetFunction.Match(mintickerdate, Range(ws.Cells(tickerstopper, 2), ws.Cells(tickerstopper - 1 + tickercount, 2)), 0)
        maxtickerdaterow = tickerstopper - 1 + Application.WorksheetFunction.Match(maxtickerdate, Range(ws.Cells(tickerstopper, 2), ws.Cells(tickerstopper - 1 + tickercount, 2)), 0)
        
        'Enter tickername in summary table
        ws.Cells(summaryrow, Tickercolumn).Value = currentticker
        
        'Calculate the yearly stock change and enter into to summary table
        yearchange = ws.Cells(maxtickerdaterow, 6) - ws.Cells(mintickerdaterow, 3)
        ws.Cells(summaryrow, Yearchangecolumn).Value = yearchange
        If yearchange > 0 Then
            ws.Cells(summaryrow, Yearchangecolumn).Interior.ColorIndex = 4
        End If
        If yearchange < 0 Then
            ws.Cells(summaryrow, Yearchangecolumn).Interior.ColorIndex = 3
        End If
        ws.Cells(summaryrow, Yearchangecolumn).Borders.Color = RGB(216, 216, 216)
        
        'Calculate the stock change as a percentage
        'Some test values were 0, so an if statement is used to avoid an overflow error
        If ws.Cells(mintickerdaterow, 3) = 0 Then
            ws.Cells(summaryrow, Percentchangecolumn).Value = "Undefined"
        Else
            percentchange = (ws.Cells(maxtickerdaterow, 6) - ws.Cells(mintickerdaterow, 3)) / ws.Cells(mintickerdaterow, 3)
            ws.Cells(summaryrow, Percentchangecolumn).Value = percentchange
        End If
        
        'Sum the total stock volume
        total = Application.WorksheetFunction.Sum(Range(ws.Cells(tickerstopper, 7), ws.Cells(tickerstopper - 1 + tickercount, 7)))
        ws.Cells(summaryrow, totalcolumn).Value = total
        
        'Update ticker stopper so that it's equal to the first row of the next ticker
        tickerstopper = tickerstopper + tickercount
        
    Loop 'Loops through each ticker symbol in the sheet
    
    'Complete limits table that finds tickers with the greatest increase/decrease of stock change, and the greatest volume
    'Could compact this section more, but this makes it to be more easily modified for now
    'Find the greatest % increase and assign to limits table
    Dim greatestincrease As Double
    greatestincrease = Application.WorksheetFunction.Max(Range(ws.Cells(2, Percentchangecolumn), ws.Cells(summaryrow, Percentchangecolumn)))
    ws.Cells(2, valuelimitscolumn).Value = greatestincrease
    
    'Find the row associated with greatest % increase
    Dim greatestincreaserow As Integer
    greatestincreaserow = 1 + Application.WorksheetFunction.Match(greatestincrease, Range(ws.Cells(2, Percentchangecolumn), ws.Cells(summaryrow, Percentchangecolumn)), 0)
    
    'Assign ticker to greatest % increase row in limits table
    ws.Cells(2, tickerlimitscolumn).Value = ws.Cells(greatestincreaserow, Tickercolumn)
    
    'Find the greatest % decrease and assign to limits table
    Dim greatestdecrease As Double
    greatestdecrease = Application.WorksheetFunction.Min(Range(ws.Cells(2, Percentchangecolumn), ws.Cells(summaryrow, Percentchangecolumn)))
    ws.Cells(3, valuelimitscolumn).Value = greatestdecrease
    
    'Find the row associated with greatest % decrease
    Dim greatestdecreaserow As Integer
    greatestdecreaserow = 1 + Application.WorksheetFunction.Match(greatestdecrease, Range(ws.Cells(2, Percentchangecolumn), ws.Cells(summaryrow, Percentchangecolumn)), 0)
    
    'Assign ticker to greatest % decrease row in limits table
    ws.Cells(3, tickerlimitscolumn).Value = ws.Cells(greatestdecreaserow, Tickercolumn)
    
    'Find the greatest total volume and assign to limits table
    Dim greatestvolume As Double
    greatestvolume = Application.WorksheetFunction.Max(Range(ws.Cells(2, totalcolumn), ws.Cells(summaryrow, totalcolumn)))
    ws.Cells(4, valuelimitscolumn).Value = greatestvolume
    
    'Find the row associated with greatest % decrease
    Dim greatestvolumerow As Integer
    greatestvolumerow = 1 + Application.WorksheetFunction.Match(greatestvolume, Range(ws.Cells(2, totalcolumn), ws.Cells(summaryrow, totalcolumn)), 0)
    
    'Assign ticker to greatest % decrease row in limits table
    ws.Cells(4, tickerlimitscolumn).Value = ws.Cells(greatestvolumerow, Tickercolumn)
    
Next ws

'Turning speed slowing functions back on
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True



End Sub
