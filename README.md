# VBA-Challenge
Sub stocks()

'Set headers for the answers
    Range("I1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("K").NumberFormat = "0.00%"
    
'Set variables for holding Ticker symbol, Open price, Close price and Total volume
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Counter As Long
    Dim Total_Volume As Double

        Total_Volume = 0

        Counter = 2

        Open_Price = Cells(Counter, 3).Value

'Find last row in column A
    Last_Row = Cells(Rows.Count, "A").End(xlUp).Row

' Area to hold answers
    Dim Answers As Double
    Answers = 2

'Loop through all Ticker Symbols from A2 to the last row
    For i = 2 To Last_Row
        'Find the first cell in column A that has a different ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Put the ticker symbol in column I
                Ticker = Cells(i, 1).Value
                Range("I" & Answers).Value = Ticker
            
                Total_Volume = Total_Volume + Cells(i, 7).Value
            
                Counter = i + 1
            
                Close_Price = Cells(i, 6).Value
                
                    If Open_Price <> 0 Then
            
                    Range("J" & Answers).Value = Close_Price - Open_Price
                    Range("K" & Answers).Value = (Close_Price - Open_Price) / Open_Price
                    Range("L" & Answers).Value = Total_Volume
                    Open_Price = Cells(Counter, 3).Value
                     
                'Conditional formatting
                    Range("J" & Answers).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                    Formula1:="=0"
                    Range("J" & Answers).FormatConditions(1).Interior.ColorIndex = 4
                    Range("J" & Answers).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                    Formula1:="=0"
                    Range("J" & Answers).FormatConditions(2).Interior.ColorIndex = 3
    
                    Answers = Answers + 1
                    Total_Volume = 0
            
                    Else
                    Range("J" & Answers).Value = Close_Price - Open_Price
                    Range("K" & Answers).Value = 0
                    Range("L" & Answers).Value = Total_Volume
                    Open_Price = Cells(Counter, 3).Value
                
                'Conditional formatting
                    Range("J" & Answers).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                    Formula1:="=0"
                    Range("J" & Answers).FormatConditions(1).Interior.ColorIndex = 4
                    Range("J" & Answers).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                    Formula1:="=0"
                    Range("J" & Answers).FormatConditions(2).Interior.ColorIndex = 3
                    
                    Answers = Answers + 1
                    Total_Volume = 0
            
                End If
           
           Else
                Total_Volume = Total_Volume + Cells(i, 7).Value
                
        End If
    
    Next i

End Sub
