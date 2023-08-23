Attribute VB_Name = "Module1"
Sub YearlyStockScanner()
    
    'create variables
    Dim worksht As Worksheet
    Dim tick As String
    Dim yearlyChange As Double
    Dim percChange As Double
    Dim total As LongLong
    
    Dim greatTot As Double
    
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim customIndex As Integer
    Dim countTick As Integer
    Dim tickNum As Integer
    
    'loop through pages to write it
    For Each worksht In Worksheets
        'Write headers for info gathered
        worksht.Range("I1").Value = "Ticker"
        worksht.Range("J1").Value = "Yearly Change"
        worksht.Range("K1").Value = "Percentage Change"
        worksht.Range("L1").Value = "Total Stock Volume"
        worksht.Range("P1").Value = "Ticker"
        worksht.Range("Q1").Value = "Value"
        worksht.Range("O2").Value = "Greatest % Increase"
        worksht.Range("O3").Value = "Greatest % Decrease"
        worksht.Range("O4").Value = "Greatest Total Volume"
        
        ' get number of data rows
        lastRow = worksht.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' set variables to initial setup
        customIndex = 1
        countTick = 0
        
        'loop through data
        For i = 3 To lastRow
            If worksht.Cells(i, 1) = worksht.Cells(i - 1, 1) Then
                countTick = countTick + 1
                tickNum = countTick + 1
            Else
                ' fill in rows
                customIndex = customIndex + 1
                openPrice = worksht.Cells(i - tickNum, 3).Value
                closePrice = worksht.Cells(i - 1, 6).Value
                
                ' ticker value
                tick = worksht.Cells(i - 1, 1).Value
                worksht.Cells(customIndex, 9).Value = tick
                ' yearly change
                yearlyChange = closePrice - openPrice
                worksht.Cells(customIndex, 10).Value = yearlyChange
                ' percent change
                percChange = (closePrice / openPrice) - 1
                worksht.Cells(customIndex, 11).Value = FormatPercent(percChange)

                ' total volume
                worksht.Cells(customIndex, 12).Value = Application.WorksheetFunction.Sum(worksht.Range(worksht.Cells(i - 1, 7), worksht.Cells(i - tickNum, 7)))
                total = worksht.Cells(customIndex, 12).Value
                'find greatest total volume
                If greatTot < total Then
                    greatTot = total
                End If
                
                ' reset ticker count
                countTick = 0
            End If
        
        ' MsgBox ("One Tick Complete")
        Next i
            
        'format percent change and yearly change with colors
        lrColor = worksht.Cells(Rows.Count, "I").End(xlUp).Row
        For x = 2 To lrColor
            If worksht.Cells(x, 10).Value > 0 Then
                worksht.Cells(x, 10).Interior.Color = vbGreen
            Else
                worksht.Cells(x, 10).Interior.Color = vbRed
            End If
            
            If worksht.Cells(x, 11).Value > 0 Then
                worksht.Cells(x, 11).Interior.Color = vbGreen
            Else
                worksht.Cells(x, 11).Interior.Color = vbRed
            End If
        Next x
        
        ' input greatest increase, decrease, total
        worksht.Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(worksht.Range("K" & 2 & ":" & "K" & lrColor)))
        worksht.Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(worksht.Range("K" & 2 & ":" & "K" & lrColor)))
        worksht.Cells(4, 17).Value = Application.WorksheetFunction.Max(worksht.Range("L" & 2 & ":" & "L" & lrColor))
        
        ' find respective ticker name for ^
        For y = 2 To lrColor
            If worksht.Cells(y, 11) = worksht.Cells(2, 17).Value Then
                worksht.Cells(2, 16).Value = worksht.Cells(y, 9).Value
            ElseIf worksht.Cells(y, 11) = worksht.Cells(3, 17).Value Then
                worksht.Cells(3, 16).Value = worksht.Cells(y, 9).Value
            ElseIf worksht.Cells(y, 12) = worksht.Cells(4, 17).Value Then
                worksht.Cells(4, 16).Value = worksht.Cells(y, 9).Value
            End If
        Next y
    
    ' MsgBox ("Page Completed")
    Next worksht
    'end loop
    MsgBox ("All Pages Completed")
End Sub
