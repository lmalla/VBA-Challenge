Attribute VB_Name = "Module1"
Sub Stock_loops()
    'variable declaration
    Dim ticker As String
    Dim yearly_change As Double
    Dim opening_price As Double
    Dim closing_price As Double
    Dim percent_change As Double
    Dim total_Stock_volume As Double
    Dim no_of_rows As Long
    Dim number_tickerss As Integer
    
    
    For Each ws In Worksheets
        ws.Activate
        
        no_of_rows = Range("A1").End(xlDown).Row
        'MsgBox (no_of_rows)
        
        'Header
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Initialize variables for each worksheet.
        number_tickers = 0
        ticker = ""
        yearly_change = 0
        opening_price = 0
        percent_change = 0
        total_Stock_volume = 0
        

        For i = 2 To no_of_rows
            ticker = Cells(i, 1).Value
            
            If opening_price = 0 Then
                opening_price = Cells(i, 3).Value
            End If
            
            total_Stock_volume = total_Stock_volume + Cells(i, 7).Value
            
            'If different ticker on the list
            If Cells(i + 1, 1).Value <> ticker Then
                number_tickers = number_tickers + 1
                Cells(number_tickers + 1, 9) = ticker
                
                closing_price = Cells(i, 6)
                
                yearly_change = closing_price - opening_price
                Cells(number_tickers + 1, 10).Value = yearly_change
                
                If yearly_change > 0 Then
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
                ElseIf yearly_change < 0 Then
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
                Else
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
                End If
                
                If opening_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = (yearly_change / opening_price)
                End If
                
                Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
                
                opening_price = 0
                
                Cells(number_tickers + 1, 12).Value = total_Stock_volume
                
                total_Stock_volume = 0
            
            End If
        Next i
    Next ws
                  
    
    
End Sub



