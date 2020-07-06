Attribute VB_Name = "Module1"
' Main subroutine that calls other sub/functions.
Sub Main()
    
    ' Range variables for autofitting columns.
    Dim derivedRange, greatestRange As Range

    ' Loop through each wkSheet.
    For Each wkSheet In Worksheets
        ' Drops table
        dropTable (wkSheet)
        
        ' Creates table and sets variable for autofitting columns.
        Set derivedRange = setDerivedTable(wkSheet)
        Set greatestRange = setGreatestTable(wkSheet)
        
        ' Fills table with tickers and then autofits table range.
        listTickers (wkSheet)
        
        ' Calculate yearly change, percent change and total volume.
        deriveFields (wkSheet)
        
        ' Conditional formatting for yearly change and percentage change.
        colorFormat (wkSheet)
        
        ' Finds and fills greatest/lowest values.
        deriveGreatestFields (wkSheet)
        
        ' I'm OCD even though OCD isn't even a verb.
        derivedRange.Columns.AutoFit
        greatestRange.Columns.AutoFit
        
    Next wkSheet

End Sub

' Function that creates the row/column headers and passes the range to main for column autofit
Function setDerivedTable(ByVal wkSheet As Worksheet) As Range

    Set tableRange = wkSheet.Range("I:L")
            
    If tableRange(1, 1).Value <> "Ticker" Then
        tableRange(1, 1).Value = "Ticker"
        tableRange(1, 2).Value = "Yearly Change"
        tableRange(1, 3).Value = "Percent Change"
        tableRange(1, 4).Value = "Total Stock Volume"
        tableRange(1, 8).Value = "Ticker"
        tableRange(1, 9).Value = "Value"
        tableRange(2, 7).Value = "Greatest % Increase"
        tableRange(3, 7).Value = "Greatest % Decrease"
        tableRange(4, 7).Value = "Greatest Total Volume"
        
        'Debug.Print "Table was set for " + wkSheet.Name
    Else
        'Debug.Print "Table is already set for " + wkSheet.Name
    End If
    
    ' Pass table range back to main.
    Set setDerivedTable = tableRange

End Function

' Function that creates the row/column headers and passes the range to main for column autofit
Function setGreatestTable(ByVal wkSheet As Worksheet) As Range

    Set tableRange = wkSheet.Range("O:Q")
            
    If tableRange(1, 2).Value <> "Ticker" Then
        tableRange(1, 2).Value = "Ticker"
        tableRange(1, 3).Value = "Value"
        tableRange(2, 1).Value = "Greatest % Increase"
        tableRange(3, 1).Value = "Greatest % Decrease"
        tableRange(4, 1).Value = "Greatest Total Volume"
        
        'Debug.Print "Table was set for " + wkSheet.Name
    Else
        'Debug.Print "Table is already set for " + wkSheet.Name
    End If
    
    ' Pass table range back to main.
    Set setGreatestTable = tableRange

End Function

' Subroutine that populates distinct ticker values.
Sub listTickers(ByVal wkSheet As Worksheet)

    Set tickerList = wkSheet.Range("A:A")
    Set tickerColumn = wkSheet.Range("I:I")
        
    ' The last filled cell in the ticker column in the dataset.
    lastTicker = wkSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' The last filled cell in the derived ticker column + 1 since column is empty when initialized.
    nextTicker = wkSheet.Cells(Rows.Count, "I").End(xlUp).Row + 1
        
    ' Iterates through cells in dataset ticket column to find the next distinct value.
    For i = 2 To lastTicker
        If tickerList(i, 1) <> tickerList(i - 1, 1) Then
            tickerColumn(nextTicker, 1).Value = tickerList(i, 1).Value
            nextTicker = nextTicker + 1
        End If
    Next i
    
    ' Clean-up
    Set tickerColumn = Nothing
    Set tickerList = Nothing
    
End Sub

' Subroutine to drops derived table upon script run.
Sub dropTable(ByVal wkSheet As Worksheet)

    wkSheet.Range("I:Q").EntireColumn.Delete
    
End Sub

' Subroutine to calculate values for derived table.
Sub deriveFields(ByVal wkSheet As Worksheet)

    Dim previousRow, openValue, closeValue, volumeTotal As Double

    ' Set ranges from dataset.
    Set tickerList = wkSheet.Range("A:A")
    Set openList = wkSheet.Range("C:C")
    Set closeList = wkSheet.Range("F:F")
    Set volumeList = wkSheet.Range("G:G")
    
    ' Set ranges for derived table columns.
    Set tickerColumn = wkSheet.Range("I:I")
    Set yearChangeColumn = wkSheet.Range("J:J")
    Set percentChangeColumn = wkSheet.Range("K:K")
    Set totalVolumeColumn = wkSheet.Range("L:L")
    
    ' Sets values for the last used cell in both dataset table and derived table.
    lastRow = wkSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lastTicker = wkSheet.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Counter variable to avoid re-iterating through entire dataset.
    previousRow = 2
    
    ' Iterates through dataset and determines open and close values by finding the next distinct ticker value.
    For i = 2 To lastTicker
        ' Initialize total volume.
        volumeTotal = 0
        
        For j = previousRow To lastRow
            If tickerColumn(i, 1) <> tickerList(j - 1, 1) Then
                openValue = openList(j, 1).Value
            End If
            
            If tickerColumn(i, 1) <> tickerList(j + 1, 1) Then
                closeValue = closeList(j, 1).Value
                previousRow = j + 1
                volumeTotal = volumeTotal + volumeList(j, 1).Value
                Exit For
            End If
                
            volumeTotal = volumeTotal + volumeList(j, 1).Value
        Next j
        
        ' Set/calculate derived columns for change and total stock volume.
        yearChangeColumn(i, 1).Value = closeValue - openValue
        
        ' If loop to prevent overflow errors from dividing by 0.
        If openValue > 0 Then
            percentChangeColumn(i, 1).Value = yearChangeColumn(i, 1) / openValue
        Else
            percentChangeColumn(i, 1).Value = 0
        End If
        
        totalVolumeColumn(i, 1).Value = volumeTotal
        
        ' Sets format for percent.
        percentChangeColumn.Cells(i, 1).NumberFormat = "0.00%"
        
    Next i
    
    ' Clean-up.
    Set tickerList = Nothing
    Set openList = Nothing
    Set closeList = Nothing
    Set tickerColumn = Nothing
    Set yearChangeColumn = Nothing
    Set percentChangeColumn = Nothing
    Set totalVolumeColumn = Nothing
    
End Sub

' Subroutine that creates conditional formatting for color coding.
Sub colorFormat(ByVal wkSheet As Worksheet)

    lastTicker = wkSheet.Cells(Rows.Count, "I").End(xlUp).Row
    
    wkSheet.Range("J2", "J" & lastTicker).FormatConditions.Add(xlCellValue, xlGreater, "=0").Interior.ColorIndex = 4
    wkSheet.Range("J2", "J" & lastTicker).FormatConditions.Add(xlCellValue, xlLess, "=0").Interior.ColorIndex = 3
    
    wkSheet.Range("K2", "K" & lastTicker).FormatConditions.Add(xlCellValue, xlGreater, "=0").Interior.ColorIndex = 4
    wkSheet.Range("K2", "K" & lastTicker).FormatConditions.Add(xlCellValue, xlLess, "=0").Interior.ColorIndex = 3
    
End Sub

' Subroutine that calculates derived values for greatest section of table.
Sub deriveGreatestFields(ByVal wkSheet As Worksheet)

    Dim greatestPercent, lowestPercent, mostVolume As Double
    Dim greatestTicker, lowestTicker, mostTicker As String
    
    ' Set ranges for derived columns
    Set tickerColumn = wkSheet.Range("I:I")
    Set percentChangeColumn = wkSheet.Range("K:K")
    Set totalVolumeColumn = wkSheet.Range("L:L")
    
    ' Set ranges for greatest cells.
    Set greatestPercentTicker = wkSheet.Range("P2")
    Set greatestPercentValue = wkSheet.Range("Q2")
    Set lowestPercentTicker = wkSheet.Range("P3")
    Set lowestPercentValue = wkSheet.Range("Q3")
    Set greatestVolumeTicker = wkSheet.Range("P4")
    Set greatestVolumeValue = wkSheet.Range("Q4")
    
    ' Sets value for last used cell in derived table.
    lastRow = wkSheet.Cells(Rows.Count, "K").End(xlUp).Row
    
    ' Initializes values for greatest section of table.
    greatestTicker = tickerColumn(2, 1).Value
    lowestTicker = tickerColumn(2, 1).Value
    mostTicker = tickerColumn(2, 1).Value
    greatestPercent = percentChangeColumn(2, 1).Value
    lowestPercent = percentChangeColumn(2, 1).Value
    mostVolume = totalVolumeColumn(2, 1).Value
    
    ' Iterates through each related column in derived table comparing values to find greatest values.
    For i = 3 To lastRow
        If percentChangeColumn(i, 1).Value > greatestPercent Then
            greatestPercent = percentChangeColumn(i, 1).Value
            greatestTicker = tickerColumn(i, 1).Value
        ElseIf percentChangeColumn(i, 1).Value < lowestPercent Then
            lowestPercent = percentChangeColumn(i, 1).Value
            lowestTicker = tickerColumn(i, 1).Value
        End If
        
        If totalVolumeColumn(i, 1).Value > mostVolume Then
            mostVolume = totalVolumeColumn(i, 1).Value
            mostTicker = tickerColumn(i, 1).Value
        End If
    Next i
    
    ' Sets values in greatest section of table.
    greatestPercentTicker.Value = greatestTicker
    greatestPercentValue.Value = greatestPercent
    lowestPercentTicker.Value = lowestTicker
    lowestPercentValue.Value = lowestPercent
    greatestVolumeTicker.Value = mostTicker
    greatestVolumeValue.Value = mostVolume
    
    ' Sets formatting for percent.
    greatestPercentValue.NumberFormat = "0.00%"
    lowestPercentValue.NumberFormat = "0.00%"
    
    ' Clean-up.
    Set tickerColumn = Nothing
    Set percentChangeColumn = Nothing
    Set totalVolumeColumn = Nothing
    Set greatestPercentTicker = Nothing
    Set greatestPercentValue = Nothing
    Set lowestPercentTicker = Nothing
    Set lowestPercentValue = Nothing
    Set greatestVolumeTicker = Nothing
    Set greatestVolumeValue = Nothing
    
End Sub
