Attribute VB_Name = "Module2"
Sub stock_market_data_B()

Cells(1, 11).Value = "Yearly Change"
Cells(1, 10).Value = "Ticker"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

'Bonus
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"
Columns("P").AutoFit


Dim ticker As String

Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

Dim Summary_Table_Row As Integer


Summary_Table_Row = 2

For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row

    If Cells(i, 2).Value = 20200102 Then
        Opening_Price = Cells(i, 3).Value
    End If
    
    If Cells(i, 2).Value = 20201231 Then
       Closing_Price = Cells(i, 6).Value
    End If
             
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        
        ticker = Cells(i, 1).Value
        Range("J" & Summary_Table_Row).Value = ticker
       
        
        Yearly_Change = Closing_Price - Opening_Price
        Range("K" & Summary_Table_Row).Value = Yearly_Change
        
        
        Percent_Change = (Yearly_Change / Opening_Price) * 100
        Range("L" & Summary_Table_Row).Value = Percent_Change
        
        
        Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Stock_Volume = 0
            
        
        Columns("K").AutoFit
        Columns("L").AutoFit
        Columns("M").AutoFit
        
     
    Else
     
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
    End If
       
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
        
    ElseIf Cells(i, 11).Value < 0 Then
        Cells(i, 11).Interior.ColorIndex = 3
     End If

Next i

'Bonus
Range("R2").Value = Application.WorksheetFunction.Max(Range("l:l"))

Range("R3").Value = Application.WorksheetFunction.Min(Range("L:L"))
      
Range("R4").Value = Application.WorksheetFunction.Max(Range("M:M"))
    
  For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
    If Cells(i, 12).Value = Range("R2").Value Then
        Range("Q2").Value = Cells(i, 10).Value
    End If
                                        
    If Cells(i, 12).Value = Range("R3").Value Then
        Range("Q3").Value = Cells(i, 10).Value
    End If
                                        
    If Cells(i, 13).Value = Range("R4").Value Then
        Range("Q4").Value = Cells(i, 10).Value
    End If

Next i

End Sub

Sub stock_market_data_C()

Cells(1, 11).Value = "Yearly Change"
Cells(1, 10).Value = "Ticker"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

'Bonus
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"
Columns("P").AutoFit


Dim ticker As String

Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

Dim Summary_Table_Row As Integer


Summary_Table_Row = 2

For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row

    If Cells(i, 2).Value = 20200102 Then
        Opening_Price = Cells(i, 3).Value
    End If
    
    If Cells(i, 2).Value = 20201231 Then
       Closing_Price = Cells(i, 6).Value
    End If
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        
        ticker = Cells(i, 1).Value
        Range("J" & Summary_Table_Row).Value = ticker
       
        
        Yearly_Change = Closing_Price - Opening_Price
        Range("K" & Summary_Table_Row).Value = Yearly_Change
        
        
        Percent_Change = (Yearly_Change / Opening_Price) * 100
        Range("L" & Summary_Table_Row).Value = Percent_Change
        
        
        Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Stock_Volume = 0
            
        
        Columns("K").AutoFit
        Columns("L").AutoFit
        Columns("M").AutoFit
        
     
    Else
     
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
    End If
       
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
        
    ElseIf Cells(i, 11).Value < 0 Then
        Cells(i, 11).Interior.ColorIndex = 3
     End If

Next i

'Bonus
Range("R2").Value = Application.WorksheetFunction.Max(Range("l:l"))

Range("R3").Value = Application.WorksheetFunction.Min(Range("L:L"))
      
Range("R4").Value = Application.WorksheetFunction.Max(Range("M:M"))
    
  For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
    If Cells(i, 12).Value = Range("R2").Value Then
        Range("Q2").Value = Cells(i, 10).Value
    End If
                                        
    If Cells(i, 12).Value = Range("R3").Value Then
        Range("Q3").Value = Cells(i, 10).Value
    End If
                                        
    If Cells(i, 13).Value = Range("R4").Value Then
        Range("Q4").Value = Cells(i, 10).Value
    End If

Next i

End Sub
