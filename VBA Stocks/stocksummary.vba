Attribute VB_Name = "Module1"
Sub stocksummary2():


For Each ws In Worksheets
    Dim ticker As String
    Dim summarytablerow As Integer
    summarytablerow = 2
    Dim totalvolume As Double
    totalvolume = 0
    Dim maxvolume As Double
    maxvolume = 0
    Dim maxpercent As Double
    maxpercent = 0
    Dim minpercent As Double
    minpercent = 0
    Dim percentchange As Double
    percentchange = 0
    Dim opening As Double
    Dim yearlychange As Double
    yearlychange = 0
 
        
    'create headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 12).Value = "Total Yearly Volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Yearly Change"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest Percent Increase"
    ws.Cells(3, 14).Value = "Greatest Percent Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    'set range for columns
    lastrow12 = ws.Cells(Rows.Count, 12).End(xlUp).Row
    lastrow10 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    lastrow11 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set first opening value
    opening = ws.Cells(2, 3).Value
    
    'to find ticker, totalvolume, opening and closing
    For i = 2 To lastrow1
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Range("I" & summarytablerow).Value = ticker
            ws.Range("L" & summarytablerow).Value = totalvolume
            closing = ws.Cells(i, 6).Value
            yearlychange = closing - opening
            percentchange = (closing - opening) / (opening + 0.01) * 100
            ws.Range("K" & summarytablerow).Value = percentchange
            ws.Range("J" & summarytablerow).Value = yearlychange
            summarytablerow = summarytablerow + 1
            totalvolume = 0
            yearlychange = 0
            opening = Cells(i + 1, 3).Value
        Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    'conditional formatting - color
    For m = 2 To lastrow10
        If ws.Cells(m, 10).Value < 0 Then
            ws.Cells(m, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(m, 10).Interior.ColorIndex = 4
            
        End If
    Next m
    
    'to find greatest total yearly volume and deposit info into side table
    For j = 2 To lastrow12
        If ws.Cells(j, 12).Value > maxvolume Then
            maxvolume = ws.Cells(j, 12).Value
            ws.Cells(4, 16).Value = maxvolume
            ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
        End If
    Next j
    
    'to find greatest percent increase, greatest percent decrease, and deposit info into side table
    For k = 2 To lastrow11
        If ws.Cells(k, 11).Value > maxpercent Then
            maxpercent = ws.Cells(k, 11).Value
            ws.Cells(2, 16).Value = maxpercent
            ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
        ElseIf ws.Cells(k, 11).Value < minpercent Then
            minpercent = ws.Cells(k, 11).Value
            ws.Cells(3, 16).Value = minpercent
            ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
        End If
    Next k
        
Next ws

End Sub




