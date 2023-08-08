'Option Explicit

Sub Stock_Data()

' Set ws as worksheet
Dim headers() As Variant
Dim ws As Worksheet
Dim wb As Workbook
Dim i As Long

Set wb = ActiveWorkbook

'Set headers

headers() = Array("Ticker", "Yearly_Change", "Percent_Change", "Stock_Volume", " ", " ", " ", "Ticker", "Value")

For Each ws In wb.Sheets
    With ws
    .Rows(1).Value = ""
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 9 + i).Value = headers(i)
    
    Next i
    .Rows(1).Font.Underline = True
    .Rows(1).VerticalAlignment = xlCenter
    End With
Next ws

    'Loop through all sheets
    For Each ws In Worksheets
    
    'Set Variables
    Dim Ticker_Sym As String
    Ticker_Sym = " "
    Dim Op_Price As Double
    Op_Price = 0
    Dim Cl_Price As Double
    Cl_Price = 0
    Dim Total_Vol As Double
    Total_Vol = 0
    Dim Price_Change As Double
    Price_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Great_Increase As Double
    Great_Increase = 0
    Dim Great_Decrease As Double
    Great_Decrease = 0
    Dim Great_Inc_Ticker_Sym As String
    Great_Inc_Ticker_Sym = " "
    Dim Great_Dec_Ticker_Sym As String
    Great_Dec_Ticker_Sym = " "
    Dim Great_Vol_Ticker_Sym As String
    Great_Vol_Ticker_Sym = " "
    Dim Great_Vol As Double
    Great_Vol = 0
    
    'Set Summary columns and table
    Dim Summary_Row As Long
    Summary_Row = 2
    Dim Lastrow As Long
    
    'Loop through all sheets to last cell
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set open prices
    Op_Price = ws.Cells(2, 3).Value
    For i = 2 To Lastrow
    
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         Ticker_Sym = ws.Cells(i, 1).Value
        
        'Calculate Price Change
         Cl_Price = ws.Cells(i, 6).Value
         Price_Change = Cl_Price - Op_Price
        
        'Calculate Percent Change
         If Op_Price <> 0 Then
         Percent_Change = (Price_Change / Op_Price) * 100
         End If
        
        'Calculate Ticker Symbol total volume
         Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
        'Place Ticker Symbol in Summary Column
         ws.Range("I" & Summary_Row).Value = Ticker_Sym
        'Place Price Change in Summary Column
         ws.Range("J" & Summary_Row).Value = Price_Change
        
        'Place Percent Change in Summary Column
         ws.Range("K" & Summary_Row).Value = (CStr(Percent_Change) & "%")
        
        'Place Total Volume in Summary Column
         ws.Range("L" & Summary_Row).Value = Total_Vol
        
        'Add color to cell based on outcome
        If (Price_Change > 0) Then
            ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
            
        ElseIf (Price_Change < 0) Then
                ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
        End If
        
        'Next Row
         Summary_Row = Summary_Row + 1
        
        '
        Op_Price = ws.Cells(i + 1, 3).Value
        
        'Calculate Greatest Volumes
        If (Percent_Change > Great_Increase) Then
            Great_Increase = Percent_Change
            Great_Inc_Ticker_Sym = Ticker_Sym
            
        ElseIf (Percent_Change < Great_Decrease) Then
            Great_Decrease = Percent_Change
            Great_Dec_Ticker_Sym = Ticker_Sym
            
        End If
        
        If (Total_Vol > Great_Vol) Then
            Great_Vol = Total_Vol
            Great_Vol_Ticker_Sym = Ticker_Sym
            
        End If
        
        'Reset Values
        Percent_Change = 0
        Total_Vol = 0
        
        
      Else
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
      End If
        
    Next i
    
        'Print Headers and values in Summary Table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P2").Value = Great_Inc_Ticker_Sym
        ws.Range("P3").Value = Great_Dec_Ticker_Sym
        ws.Range("P4").Value = Great_Vol_Ticker_Sym
        ws.Range("Q2").Value = (CStr(Great_Increase) & "%")
        ws.Range("Q3").Value = (CStr(Great_Decrease) & "%")
        ws.Range("Q4").Value = Great_Vol
        
        
    Next ws
    
       
End Sub

