Sub stock_data()
  ' --------------------------------------------
  ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
  Dim WS As Worksheet
    For Each WS In Worksheets
      
      ' Determine the Last Row
      LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

      ' Add Header for Summary Table
      WS.Cells(1, "I").Value = "Ticker"
      WS.Cells(1, "J").Value = "Yearly Change"
      WS.Cells(1, "K").Value = "Percent Change"
      WS.Cells(1, "L").Value = "Total Stock Volume"

      ' Create Variables to hold Values
      Dim Ticker As String
      Dim Open_Price As Double
      Dim Close_Price As Double
      Dim Yearly_Change As Double
      Dim Percent_Change As Double
      Dim Total_volume As Double
      Total_volume = 0
      
      ' Keep track of the location for each ticker symbol in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
      
      Dim i As Long
      
      ' Set Initial Open Price
      Open_Price = WS.Cells(2, 3).Value

      ' Loop through all ticker symbols
      For i = 2 To LastRow

        ' Check if we are still within the same ticker symbol, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
          ' Set the Ticker name
          Ticker = WS.Cells(i, 1).Value
          WS.Cells(Summary_Table_Row, "I").Value = Ticker

          ' Set Close Price
          Close_Price = WS.Cells(i, 6).Value

          ' Add Yearly Change
          Yearly_Change = Close_Price - Open_Price
          WS.Cells(Summary_Table_Row, "J").Value = Yearly_Change

          ' Add Percent Change
            If (Open_Price = 0 And Close_Price = 0) Then
                Percent_Change = 0
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                Percent_Change = 1
            Else
                Percent_Change = Yearly_Change / Open_Price
                WS.Cells(Summary_Table_Row, "K").Value = Percent_Change
                WS.Cells(Summary_Table_Row, "K").NumberFormat = "0.00%"
            End If
        
          ' Add Total Volume
          Total_volume = Total_volume + WS.Cells(i, 7).Value
          WS.Cells(Summary_Table_Row, "L").Value = Total_volume

          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1

          ' Reset the Total Volume
          Total_volume = 0

          ' Reset the Open Price
          Open_Price = WS.Cells(i + 1, 3).Value
          
        ' If the cell immediately following a row is the same ticker...
        Else
        
          Total_volume = Total_volume + WS.Cells(i, 7).Value
        
        End If
      Next i

      ' Determine the Last Row of Yearly Change 
      YCLastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row
        
        ' Set the Cell Colors
        For j = 2 To YCLastRow
            If (WS.Cells(j, 10).Value > 0 Or WS.Cells(j, 10).Value = 0) Then
                WS.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf WS.Cells(j, 10).Value < 0 Then
                WS.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

  Next WS

End Sub
