Sub stockProcessor():
    ' 'StockProcessor' is the main function that I am using to execute the functionality defined in the challenge
    ' ws is a variable that stores which ws we are currently on in the For Each loop
    Dim ws As Worksheet
    
    ' This For Each loop will loop through each worksheet in the Excel document and make all relavent modifications
    For Each ws In Worksheets
        ' These lines of code initializes all of the column names to their desired values (as defined in the challenge)
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ' This variable stores the 'greatest percent increase' value that has to be caluclated at the end of the sheet calculations
        Dim greatestPtIncVal As Double
        ' This variable stores the 'greatest percent decrease' value that has to be caluclated at the end of the sheet calculations
        Dim greatestPtDecVal As Double
        ' This variable stores the 'greatest total volume' value that has to be caluclated at the end of the sheet calculations
        Dim greatestTotalVolume As LongLong
        ' This variable stores the ticker of the company associated with the 'greatest percent increase' value
        Dim greatestPtIncTckr As String
        ' This variable stores the ticker of the company associated with the 'greatest percent decrease' value
        Dim greatestPtDecTckr As String
        ' This variable stores the ticker of the company associated with the 'greatest total volume' value
        Dim greatestTotalVolumeTicker As String
        
        ' This variable stores which rows in the raw output (columns I-L) a company's data is going - this iterates each time we start
        ' to read the data of a new company and all company output data from the current company has to be processed
        Dim outputIndex As Integer
        ' This variable stores which row in the raw input data we currently are accessing (note that this has been declared as a Long
        ' to account for how long the datasets are)
        Dim infoIndex As Long
        ' This vartiable stores the ticker of the company we are currently looking at
        Dim currentTicker As String
        
        ' These are just variable initializations that set the initial variable values for each variable defined above (for each worksheet)
        outputIndex = 2
        infoIndex = 2
        greatestPtIncVal = 0
        greatestPtDecVal = 0
        greatestTotalVolume = 0
        greatestPtIncTckr = ""
        greatestPtDecTckr = ""
        greatestTotalVolumeTicker = ""
        
        ' This while loop loops through each of the companys for the given quarter and records the appropriate values defined below
        While Not (Cells(infoIndex, 1).Value = "")
            ' This stores what the initial opening value is for a given company
            Dim initOpen As Double
            ' This stores what the final close vaue is for a given company - this is only set to an actual value at the end of the data for that company
            Dim finalClose As Double
            ' This records the total volume experienced by the given company durring the given quarter
            Dim volume As LongLong
            
            ' These statements record the basic information for the company before we actually start iterating through the company data
            currentTicker = ws.Cells(infoIndex, 1).Value
            initOpen = ws.Cells(infoIndex, 3).Value
            volume = 0
            
            ' This while loops loops through the company data to record what the total volume for that company is
            While (ws.Cells(infoIndex, 1).Value = currentTicker)
                volume = volume + ws.Cells(infoIndex, 7).Value
                
                infoIndex = infoIndex + 1
            Wend
            
            ' This records the finalClose value that is defined above to be used in calculating the assigned metrics
            finalClose = ws.Cells(infoIndex - 1, 6).Value
            
            ' This code assigns the metrics for this company to their associated output rows in the output columns
            ws.Cells(outputIndex, 9).Value = currentTicker
            ws.Cells(outputIndex, 10).Value = finalClose - initOpen
            ws.Cells(outputIndex, 11).Value = (finalClose / initOpen) - 1
            ws.Cells(outputIndex, 12).Value = volume
            
            ' This line automatically formats the 'percent change' values to be in the percent format
            ws.Cells(outputIndex, 11).NumberFormat = "0.00%"
            
            ' This conditional block will change the color of the quarterly change cell for the given company to green
            ' or red depending on whether the value stored in that cell is positive or negative
            If ws.Cells(outputIndex, 10).Value > 0 Then
                ws.Cells(outputIndex, 10).Interior.ColorIndex = 4
            ElseIf (ws.Cells(outputIndex, 10).Value < 0) Then
                ws.Cells(outputIndex, 10).Interior.ColorIndex = 3
            End If
            
            ' This block of code checks if the current company has a percent change larger than the current 'greatest % increase' value,
            ' smaller than the 'greatest % decrease' value or larger than the current 'greatest total volume' value and updates their respective
            ' values delcared above appropriately.
            If (ws.Cells(outputIndex, 11).Value > greatestPtIncVal) Then
                greatestPtIncVal = ws.Cells(outputIndex, 11).Value
                greatestPtIncTckr = currentTicker
            End If
            
            If (ws.Cells(outputIndex, 11).Value < greatestPtDecVal) Then
                greatestPtDecVal = ws.Cells(outputIndex, 11).Value
                greatestPtDecTckr = currentTicker
            End If
            
            If (ws.Cells(outputIndex, 12).Value > greatestTotalVolume) Then
                greatestTotalVolume = ws.Cells(outputIndex, 12).Value
                greatestTotalVolumeTicker = currentTicker
            End If
            
            ' The output index is iterated so that the data we record for the next company will be on the line following the current company's
            ' data in the data output
            outputIndex = outputIndex + 1
        Wend
        
        ' This will set all of the 'Greatest' values defined in the challenge for the current sheet right before we move on to the next sheet
        ws.Range("O2").Value = greatestPtIncTckr
        ws.Range("O3").Value = greatestPtDecTckr
        ws.Range("O4").Value = greatestTotalVolumeTicker
        
        ws.Range("P2").Value = greatestPtIncVal
        ws.Range("P3").Value = greatestPtDecVal
        ws.Range("P4").Value = greatestTotalVolume
        
        ' This formats the percentage values to be percentages
        ws.Range("P2:P3").NumberFormat = "0.00%"
        
    Next ws
End Sub
' This is just a helper method that I created to make it easier to clear all of the output data to run tests
Sub ClearContentsHelper()
    Dim ws As Worksheet
    Dim lastRowIndex As Integer
    
    For Each ws In Worksheets
        lastRowIndex = ws.Cells(Rows.Count, 9).End(xlUp).Row
        deleteRange = "I1:P" + CStr(lastRowIndex)
        
        ws.Range(deleteRange).Interior.ColorIndex = xlNone
        ws.Range(deleteRange).ClearContents
    Next ws
End Sub