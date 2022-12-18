Attribute VB_Name = "Module1"
Sub stock_ticker()
    ' Define terms
    Dim ticker As String
'    Dim open_, high, low, close_, change As Double
'    Dim i, last_row, summary_row As Long
'    Dim j As Integer
    Dim open_, close_, change As Double
    Dim total_stock, volume_total, greatest_vol As Double
    Dim open_row, percent, greatest, lowest As Double

    
    
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    
    
    
    ' Find last row and check it
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    open_row = 2
    
    open_ = 0
    close_ = 0
    volume_total = 0
    greatest = 0
    lowest = 0
    greatest_vol = 0
    
'    MsgBox (last_row)
    
' change / open = percent


    For i = 2 To last_row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker = Cells(i, 1).Value
            open_ = open_ + Cells(open_row, 3).Value
            close_ = close_ + Cells(i, 6).Value
            
            change = close_ - open_
            percent = change / open_
            
             volume_total = volume_total + Cells(i, 7).Value
            
            Range("M" & summary_row).Value = volume_total
            
            Range("J" & summary_row).Value = ticker
            Range("K" & summary_row).Value = change
            
            Range("L" & summary_row).Value = FormatPercent(percent)
            
            
            
            summary_row = summary_row + 1
            open_row = i + 1
            
            open_ = 0
            close_ = 0
            change = 0
            volume_total = 0
            
'           ---------------------------------------
            
           
        Else
            volume_total = volume_total + Cells(i, 7).Value
'            open_ = open_ + Cells(i, 4).Value
'            close_ = close_ + Cells(i, 5).Value
'            change = open_ - close_

        
        
'        ticker = Cells(i, 1).Value
        
    End If
    
Next i

'    summary_row = 0
'
'    For i = 2 To last_row
'
'        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
''
''            volume_total = volume_total + Cells(i, 7).Value
''
''            Range("M" & summary_row).Value = volume_total
''
''            summary_row = summary_row + 1
''            volume_total = 0
'
''        Else
''            volume_total = volume_total + Cells(i, 7).Value
'
'    End If
'
'Next i

    For i = 2 To last_row
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
            
        Else
            Cells(i, 11).Interior.ColorIndex = 3
    
    End If
Next i


For i = 2 To last_row
    row_num = 2
    If Cells(i, 12).Value > greatest Then
      greatest = Cells(i, 12)
     Cells(row_num, 18).Value = FormatPercent(greatest)
     Cells(row_num, 17).Value = Cells(i, 10).Value
     
    End If
Next i

For i = 2 To last_row
    row_num = 3
    If Cells(i, 12).Value < lowest Then
      lowest = Cells(i, 12)
     Cells(row_num, 18).Value = FormatPercent(lowest)
     Cells(row_num, 17).Value = Cells(i, 10).Value
     
    End If
Next i


For i = 2 To last_row
    row_num = 4
    If Cells(i, 13).Value > greatest_vol Then
      greatest_vol = Cells(i, 13)
     Cells(row_num, 18).Value = greatest_vol
     Cells(row_num, 17).Value = Cells(i, 10).Value
     
    End If
Next i


'    For i = 2 To last_row
'        If Cells(i + 1, 12).Value > Cells(i, 12).Value Then
'        greatest = Cells(i + 1, 12).Value
'        Range("Q2").Value = greatest
'
'    End If
'Next i


'greatest = WorksheetFunction.Max(percent)
'Range("Q2").Value = greatest

        
End Sub

