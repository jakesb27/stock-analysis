Sub StockAnalysis()

    Dim i, wsCount As Integer
    
    wsCount = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To wsCount
        ActiveWorkbook.Sheets(i).Select
        AddTemplate
        FillTemplate
        PolarValues
        FormatSheet
        Range("I1:Q4").Columns.AutoFit
        Range("A1").Select
    Next i
    
    ActiveWorkbook.Sheets(1).Select

End Sub

Private Sub FormatSheet()
' Apply number and color formats
    
    Dim i, rNum As Long
    
    rNum = Range("J1").End(xlDown).Row
    
    For i = 2 To rNum
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = vbGreen
        Else
            Cells(i, 10).Interior.Color = vbRed
        End If
    Next i
    
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00%"
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    Range("Q4").Select
    Selection.NumberFormat = "0"

End Sub

Private Sub FillTemplate()
' Populates column I with the unique ticker names from column A

    Dim i, t, v, rNum As Long
    Dim oVal, cVal As Double
    Dim ticker As String
    Dim stkVol As LongLong
    
    t = 1
    rNum = Range("A1").End(xlDown).Row
    stkVol = 0
    
    ' Loop through all cells in row A
    For i = 2 To rNum
        ticker = Cells(t, 9).Value
        
        ' Check if value in A is different than value in I
        If Cells(i, 1).Value <> ticker Then
        
            ' If not on first row update column I with stats
            If i > 2 Then
                Cells(t, 10).Value = cVal - oVal
                Cells(t, 11).Value = (cVal - oVal) / oVal
                Cells(t, 12).Value = stkVol
                stkVol = 0
            End If
            
            t = t + 1
            ' Set opening value and set stock volume
            Cells(t, 9).Value = Cells(i, 1).Value
            oVal = Cells(i, 3).Value
            stkVol = stkVol + Cells(i, 7).Value
        
        ' Value in A matches value in I
        Else
        
            ' If on the last row
            If i = rNum Then
                Cells(t, 10).Value = Cells(i, 6).Value - oVal
                Cells(t, 11).Value = (Cells(i, 6).Value - oVal) / oVal
            End If
            
            ' Overwrite close value and add stock volume
            cVal = Cells(i, 6).Value
            stkVol = stkVol + Cells(i, 7).Value
        End If
    Next i

End Sub

Private Sub PolarValues()
' Calculates the greatest % increase, % decrease and stock volume

    Dim i, rNum As Long
    Dim pctInc, pctDec As Double
    Dim stkVol As LongLong
    
    rNum = Range("I2").End(xlDown).Row
    pctInc = Cells(2, 11).Value
    pctDec = Cells(2, 11).Value
    stkVol = Cells(2, 12).Value
    
    For i = 2 To rNum
        If Cells(i, 11).Value >= pctInc Then
            pctInc = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = pctInc
        End If
        
        If Cells(i, 11).Value <= pctDec Then
            pctDec = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = pctDec
        End If
        
        If Cells(i, 12).Value >= stkVol Then
            stkVol = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = stkVol
        End If
    Next i

End Sub

Private Sub AddTemplate()
' Adds metadata template

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

End Sub
