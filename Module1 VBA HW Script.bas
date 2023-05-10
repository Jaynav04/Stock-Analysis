Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_Analysis()

Dim ws As Worksheet

For Each ws In Worksheets


'Declare and set worksheet
Dim WorksheetName As String
WorksheetName = ws.Name

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable and set to first row
Dim Ticker As Long
Ticker = 2

'Created another variable for start of ticker block this will start on row 2
Dim j As Long
j = 2

'Define i variable
Dim i As Long

'Define Lastrow of column A
Dim LastRowA As Long
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Do loop of current worksheet to Lastrow
For i = 2 To LastRowA

    'Ticker symbol output in column I
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value

        'Calculate and write Yearly Change in column J (#10)
        ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    
        'Conditional formating
            If ws.Cells(Ticker, 10).Value < 0 Then
                
        'Set cell background color to red
            ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
            Else
                
        'Set cell background color to green
            ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
            End If
            'Calculate and write percent change in column K
            'Variable for percent change calculation
                 Dim PerChange As Double
        
                If ws.Cells(i, 3).Value <> 0 Then
                PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                'Percent formating
                ws.Cells(Ticker, 11).Value = Format(PerChange, "Percent")
                    
                 Else
                    
                ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                    
                End If
         
    'Calculate and write total volume in column L
        ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
       'Increase TickCount by 1
                Ticker = Ticker + 1
                
                'Set new start row of the ticker block
                j = i + 1

    End If

    Next i
       'last row column I
        Dim LastRowI As Long
        'Variable for greatest increase calculation
        Dim GreatIncr As Double
        'Variable for greatest decrease calculation
        Dim GreatDecr As Double
        'Variable for greatest total volume
        Dim GreatVol As Double
        
        'Prepare for summary
     
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            'Loop for summary percentages
            For i = 2 To LastRowI
            
                'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i

    Next ws

End Sub
