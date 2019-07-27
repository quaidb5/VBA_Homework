Sub VBA_Homework()


For Each ws In Worksheets
    ws.Activate
    
    'Declare Variables'
    Dim i As Long
    Dim Ticker As String
    Dim TotalVolume As LongLong
    Dim YearChange As Double
    Dim YearPercent As Double
    Dim OpenP As Double
    Dim CloseP As Double
    Dim TableRow As Integer
    Dim RowCount As Long

    'Initializing Specific Variables'
    RowCount = Application.ActiveSheet.UsedRange.Rows.Count
    TotalVolume = 0
    Sheetname = ws.Name
    TableRow = 2

    'Apply titles to the data table'
    Range("I1").Value = "Stock Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

'---------------------------------------------------------------'

    'Calculating the Total Volume for each stock'
    For i = 2 To RowCount

        'When stock ticker's mismatch begin if'
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Establish Stock Ticker for the Table
            Ticker = Cells(i, 1).Value

            'Taking Total Volume from all iterations of a stock, adding the last one to a new variable'
            Ticker_Total = TotalVolume + Cells(i, 7).Value

            'Printing the stock ticker and total volume on the table'
            Range("I" & TableRow).Value = Ticker
            Range("L" & TableRow).Value = Ticker_Total

            'Iterate variable to move onto the next row of the table'
            TableRow = TableRow + 1

            'Reset Total Volume variables'
            Ticker_Total = 0
            TotalVolume = 0
        
        'When the stock symbols match, take the volume and add it to the current total'
        Else
            TotalVolume = TotalVolume + Cells(i, 7).Value

        End If

    
    Next i

'---------------------------------------------------------------------'

    'Reset the table row to begin at the top'
    TableRow = 2
    
    'Set open price to first ticker as a range'
    OpenP = Range("C2")

    'Loop to Generate Yearly Change and Percent Change'
    For i = 2 To RowCount
        
        'If top cell is bigger than next cell'
        If Cells(i, 2).Value > Cells(i + 1, 2).Value Then

            'Grab closing price for ticker of top cell'
            CloseP = Cells(i, 6).Value

            'Calculate raw change in price'
            YearChange = CloseP - OpenP
            
            'Using If statement because page P has an all zero security that blew up the program'
            If OpenP = 0 Then
            
                'YearPercent = Str("N/A")
                YearPercent = 0
                
            Else
                'Calculate percent change in price'
                YearPercent = YearChange / OpenP
                
            End If
                
            'Place Values into summary table'
            Range("J" & TableRow).Value = YearChange
            
                
           'Else
            Range("K" & TableRow).Value = YearPercent

            'End If

            'Zero out relevant variables'
            YearChange = 0
            YearPercent = 0
            OpenP = 0
            CloseP = 0

            'Iterate Table Rows'
            TableRow = TableRow + 1

            'opening price for the next stock'
            OpenP = Cells(i + 1, 3)

        End If
            
    Next i

'-------------------------------------------------------------------'

    'Conditional Formating for Total Yearly Change'
    For i = 2 To RowCount

        'Cells greater than 0 are green'
        If Cells(i, 10) > 0 Then
            
            Cells(i, 10).Interior.ColorIndex = 4
            
        End If
        
        'Cells less than 0 are red'
        If Cells(i, 10) < 0 Then
            
           Cells(i, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i

'-------------------------------------------------------------------'

    'Print side table'
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"

    'Initialize Variables'
    HighPercent = 0

    'Loop for finding Greatest % Increase'
    For i = 2 To RowCount 'RowCount is already declared--make sure it gets reset to 0'
        If Cells(i, 11).Value > HighPercent Then
            HighPercent = Cells(i, 11).Value
            Ticker = Cells(i, 9).Value
        End If
    Next i

    'Printing Table Values'
    Range("O2") = Ticker
    Range("P2") = HighPercent

    'Initialize Variable'
    LowPercent = 0

    'Loop for finding Greates % Decrease'
    For i = 2 To RowCount
        If Cells(i, 11).Value < LowPercent Then
            LowPercent = Cells(i, 11).Value
            Ticker = Cells(i, 9).Value
        End If
    Next i

    'Printing Table Values'
    Range("O3") = Ticker
    Range("P3") = LowPercent

    'Initialize Variable
    HighVolume = 0

    'Loop for finding highest total volume'
    For i = 2 To RowCount
        If Cells(i, 12) > HighVolume Then
            HighVolume = Cells(i, 12).Value
            Ticker = Cells(i, 9).Value
        End If
    Next i

    'Printing the table values'
    Range("O4") = Ticker
    Range("P4") = HighVolume
    
'------------------------------------------------------------------'
    
    'Loop to format percentage change column and in the Best/Worst Percentage Changed'
    For i = 2 To TableRow
        
        Cells(i, 11).NumberFormat = "0.00%"
    Next i

    Cells(2, 16).NumberFormat = "0.00%"
    Cells(3, 16).NumberFormat = "0.00%"
        
    'Format columns to fit Titles and information'
    Range("A:P").Columns.AutoFit
    
    Range("A1:P1").HorizontalAlignment = xlCenter

Next ws


End Sub