Attribute VB_Name = "Module1"
Sub HW2():

For Each St In Worksheets

St.Range("J1").Value = "Ticker"
St.Range("K1").Value = "Yearly Change"
St.Range("L1").Value = "Percent Change"
St.Range("M1").Value = "Total Stock Volume"

LastRow = St.Cells(Rows.Count, 1).End(xlUp).Row

 ' Set an initial variable for holding the brand name
  Dim TickerName As String

  ' Set an initial variable for holding the total per credit card brand
  Dim TotalStockVolume As Double
  TotalStockVolume = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Dim OpenPrice As Double
  Dim ClosePrice As Double
  
  OpenPrice = -100
  ClosePrice = -100
  
  Dim YearlyChange As Double
  Dim PercentChange As Double
  
  Dim PercentchangeNA As String
  

  ' Loop through all credit card purchases
  For i = 2 To LastRow
  
  If OpenPrice = -100 Then
    OpenPrice = St.Cells(i, 3).Value
  End If

 ClosePrice = St.Cells(i, 6).Value
 

    ' Check if we are still within the same credit card brand, if it is not...
    If St.Cells(i + 1, 1).Value <> St.Cells(i, 1).Value Then

      ' Set the Brand name
       TickerName = St.Cells(i, 1).Value

      ' Add to the Brand Total
      TotalStockVolume = TotalStockVolume + St.Cells(i, 7).Value
      
      YearlyChange = ClosePrice - OpenPrice
      
      If OpenPrice = 0 And ClosePrice = 0 Then
      PercentChange = 0
      'Range("L" & Summary_Table_Row).Value = PercentChange
      ElseIf OpenPrice = 0 And ClosePrice <> 0 Then
      PercentchangeNA = "NA"
      'Range("L" & Summary_Table_Row).Value = PercentchangeNA
      Else
      PercentChange = (ClosePrice - OpenPrice) / OpenPrice
      'Range("L" & Summary_Table_Row).Value = PercentChange
      End If
      
    

      ' Print the Credit Card Brand in the Summary Table
      St.Range("J" & Summary_Table_Row).Value = TickerName

      ' Print the Brand Amount to the Summary Table
      St.Range("M" & Summary_Table_Row).Value = TotalStockVolume
      
      St.Range("K" & Summary_Table_Row).Value = YearlyChange
      
      If St.Range("K" & Summary_Table_Row).Value > 0 Then
      St.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      
      Else
      St.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      
      If OpenPrice = 0 And ClosePrice = 0 Then
      St.Range("L" & Summary_Table_Row).Value = PercentChange
      ElseIf OpenPrice = 0 And ClosePrice <> 0 Then
      St.Range("L" & Summary_Table_Row).Value = PercentchangeNA
      Else
      St.Range("L" & Summary_Table_Row).Value = PercentChange
      End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
       TotalStockVolume = 0
       
        OpenPrice = -100
        ClosePrice = -100

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      TotalStockVolume = TotalStockVolume + St.Cells(i, 7).Value
      

    End If

  Next i
  
  Next

End Sub
