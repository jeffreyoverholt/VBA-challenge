Sub worksheets()
    

    'Set an initial variable for holding the ticker
    'referenced Credit Card Checker to add up volume total
    Dim ticker As String

    'set an initial variable for holding the total volume per ticker
    Dim ticker_volume As Double
    ticker_volume = 0
    
    'set year open
    Dim year_open As Double
    
    'set year close
    Dim year_close As Double
    
    'set value for change
    Dim yearly_change As Double
    
    'set value for top
    'https://nuvirtdatapt1-ice5461.slack.com/archives/C02SR0D8R35
    Dim top As Long
    
    'set value for % change
    Dim percent_change As String
    
    'set ws variable
    'https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
    Dim ws As Worksheet
    
    'loop through all sheets
    'https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
    'https://excelchamps.com/vba/loop-sheets/
    For Each ws In ThisWorkbook.worksheets
    ws.Activate
    
    'add summary table headers
    'https://www.mrexcel.com/board/threads/add-column-headers-in-a-worksheet-using-vba.1078803/
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Volume"
        
    'Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    
        'Loop through all tickers
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        'Check if we are still within the same ticker, if it is not...
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Set the ticker name
        ticker = Cells(i, 1).Value
      
        ' Add to the total volume
        ticker_volume = ticker_volume + Cells(i + 1, 7).Value
      
        'find where ticker starts and stops
        'https://nuvirtdatapt1-ice5461.slack.com/archives/C02SR0D8R35
        top = Range("A:A").Find(what:=Cells(i, 1).Value).Row
      
        'locate year open
        year_open = Cells(top, 3).Value
      
        'locate year closed
        year_close = Cells(i, 6).Value
      
        'calculate year change
        yearly_change = year_close - year_open
        
        'End If
              
            'correct runtime error 11 division by 0
            'https://stackoverflow.com/questions/38246478/how-to-solve-runtime-error-11-division-by-0-in-vba
            If year_open = 0 Or IsEmpty(year_open) Then
            percent_change = 0
            Else
        
        
            'calculate % change
            'print as percent
            'https://www.excelfunctions.net/vba-formatpercent-function.html
            percent_change = FormatPercent(yearly_change / year_open)
            'percent_change = yearly_change / year_open
        
            End If
        
      ' Add to the total volume
        ticker_volume = ticker_volume + Cells(i, 7).Value
        
        'Print the ticker in the Summary Table
        Range("J" & Summary_Table_Row).Value = ticker

        'Print the Brand Amount to the Summary Table
        Range("M" & Summary_Table_Row).Value = ticker_volume
      
        'print the yearly change to the summary table
         Range("K" & Summary_Table_Row).Value = yearly_change
      
        'print the % change to the summary table
        Range("L" & Summary_Table_Row).Value = percent_change
      

        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        'Reset the ticker volume
        ticker_volume = 0
        
        Else
      
        'Add to the ticker volume
        ticker_volume = ticker_volume + Cells(i, 7).Value
      
        'if cell immediately following row is the same ticker
        'Else
    
        
        End If
        
        
    Next i
        
        
    'Loop through all yearly changes to apply color
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
    
        'color code the yearly change row green if positive and red if negative
        'referenced student grade book activity
        If ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
    
        End If
    
    Next i

    Next ws
    

End Sub


