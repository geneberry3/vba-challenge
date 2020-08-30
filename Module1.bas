Attribute VB_Name = "Module1"

Sub summary_info()

    Dim Ticker_Name As String
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 1
    Dim Ticker_Total, Yearly_Open, Yearly_Close As Double
    Ticker_Total = 0
    Yearly_Open = 0
    Yearly_Close = 0
    Year_Change = 0

  For i = 2 To 800000
    
    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    Ticker_Name = Cells(i, 1).Value
    Ticker_Total = Ticker_Total + Cells(i, 6).Value
    Yearly_Close = Cells(i, 5).Value
    

    
    Range("J" & Summary_Table_Row + 1).Value = Ticker_Name
    Range("M" & Summary_Table_Row + 1).Value = Ticker_Total
    Range("L" & Summary_Table_Row + 1).Value = Yearly_Open
    Range("K" & Summary_Table_Row + 1).Value = Yearly_Close


    ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
  
    Ticker_Total = 0
  
    Else
    
    Yearly_Open = Cells(i, 3)

    Ticker_Total = Ticker_Total + Cells(i, 6).Value

        End If

    
    Next i

End Sub
