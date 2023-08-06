Attribute VB_Name = "Module1"
Sub Ticker()

  ' Set an initial variable for holding the brand name
  Dim Tracker_Name As String
  
  Dim lastrow As Long
  
  Dim Year As Integer
  
  Dim Year_Open As Double
  Dim Year_Close As Double

  ' Set an initial variable for holding the total stock volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

    'location for each ticker in the summary table
  Dim Summary_Table_Row As Long
  For Year = 2018 To 2020
      Summary_Table_Row = 2

      Worksheets(CStr(Year)).Select
  
      Cells(1, "I").Value = "Ticker"
      Cells(1, "J").Value = "Yearly_Change"
      Cells(1, "K").Value = "Percent_Change"
      Cells(1, "L").Value = "Total_Stock_Volume"
            
      ' Loop through all ticker transactions
      lastrow = Cells(Rows.Count, 1).End(xlUp).Row
      For i = 2 To lastrow
    
        ' Last row for the current ticker.
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          ' Set the ticker name
          Ticker_Name = Cells(i, 1).Value
    
          ' Year Close Value
          Year_Close = Cells(i, 6).Value
    
          ' Add to the total stock volume
          Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
          ' Print the ticker name in the Summary Table
          Range("I" & Summary_Table_Row).Value = Ticker_Name
    
          ' Print the Brand Amount to the Summary Table
          Range("J" & Summary_Table_Row).Value = Year_Close - Year_Open
    
          Range("K" & Summary_Table_Row).Value = 100 * (Year_Close - Year_Open) / Year_Open
     
          Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
          
          If (Year_Close - Year_Open) <= 0 Then
            Cells(Summary_Table_Row, 10).Interior.Color = vbRed
          Else
            Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
          End If

          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Brand Total
          Brand_Total = 0
    
        ' First row for the current ticker.
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
          ' Add to the Total_Stock_Volume
          Total_Stock_Volume = Cells(i, 7).Value
          
          ' Year Open Value
          Year_Open = Cells(i, 3).Value
        
        Else
    
            'Add to the total stock volume
          Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
          'Debug.Print Cells(i, 7).Value
          
        End If
    
      Next i
      
  Next Year
      
  MsgBox ("DONE")

End Sub





Public Sub Total()
 

End Sub
