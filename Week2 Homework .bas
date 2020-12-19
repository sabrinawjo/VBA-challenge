Attribute VB_Name = "Module1"
Sub stock()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim LastRow As Long
    Dim Ticker As String
    Dim TotalVol As Double
    Dim boxrow As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim counttime As Double
    Dim percentchange As Double
    
    
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly change"
    Cells(1, "K").Value = "Percent change"
    Cells(1, "L").Value = "Total Volume"
    ws.Columns("I:L").AutoFit
    
    boxrow = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalVol = 0
    counttime = 1
   
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            TotalVol = TotalVol + Cells(i, 7).Value
            Range("I" & boxrow).Value = Ticker
            Range("L" & boxrow).Value = TotalVol
            
            openprice = Cells(i - counttime + 1, 3).Value
            closeprice = Cells(i, 6).Value
            yearlychange = closeprice - openprice
            Range("J" & boxrow).Value = yearlychange
                If yearlychange < 0 Then
                    Range("J" & boxrow).Interior.ColorIndex = 3
                Else
                    Range("J" & boxrow).Interior.ColorIndex = 4
                End If
                
            percentchange = (closeprice - openprice) / openprice
            Range("K" & boxrow).Value = percentchange
            ws.Range("K2:" & "K" & LastRow).Style = "Percent"
            
            boxrow = boxrow + 1
            TotalVol = 0
            counttime = 1
        Else
            TotalVol = TotalVol + Cells(i, 7).Value
            counttime = counttime + 1
            
        End If
    Next i
    

    
End Sub

