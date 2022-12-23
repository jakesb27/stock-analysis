Sub StockAnalysis()

    Application.ScreenUpdating = False

    Dim i, wsCount As Integer
    
    wsCount = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To wsCount
        ActiveWorkbook.Sheets(i).Select
        AddTemplate
        SummarizeData
        PolarValues
        FormatSheet
    Next i
    
    ActiveWorkbook.Sheets(1).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Analysis has completed!", vbInformation, "Complete"

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

Private Sub SummarizeData()
' Summarizes the stock data and fills out template

    Dim i, t, rowNum As Long
    Dim oVal As Double
    Dim stkVol As LongLong

    rowNum = Range("A1").End(xlDown).Row
    t = 1

    For i = 2 To rowNum
    
        ' If on first instance of ticker name
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            t = t + 1
            Cells(t, 9).Value = Cells(i, 1).Value
            oVal = Cells(i, 3).Value
            stkVol = Cells(i, 7).Value
            
        ' If on last instance of ticker name
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(t, 10).Value = Cells(i, 6).Value - oVal
            Cells(t, 11).Value = (Cells(i, 6).Value - oVal) / oVal
            Cells(t, 12).Value = Cells(i, 7).Value + stkVol
        
        ' not first or last instance
        Else
            stkVol = stkVol + Cells(i, 7).Value
            
        End If
    Next i

End Sub

Private Sub PolarValues()
' Calculates the greatest % increase, % decrease and stock volume

    Dim i, rowNum As Integer
    Dim pctInc, pctDec As Double
    Dim stkVol As LongLong
    
    rowNum = Range("I2").End(xlDown).Row
    pctInc = Cells(2, 11).Value
    pctDec = Cells(2, 11).Value
    stkVol = Cells(2, 12).Value
    
    For i = 2 To rowNum
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

Private Sub FormatSheet()
' Apply number and color formats
    
    Dim i, rowNum As Integer
    
    rowNum = Range("J1").End(xlDown).Row
    
    For i = 2 To rowNum
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
    
    Range("I1:Q4").Columns.AutoFit
    Range("A1").Select

End Sub
