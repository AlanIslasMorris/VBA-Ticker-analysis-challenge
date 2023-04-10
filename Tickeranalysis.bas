Attribute VB_Name = "Módulo3"
Sub TickerAnalysis_2()
Dim ws As Worksheet
    Dim lastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim LastClosePrice As Double
    Dim TickerCount As Long
    Dim i As Long
    Dim Brand_Total As Double
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    For Each ws In ThisWorkbook.Worksheets
        Brand_Total = 0
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Open Price"
        ws.Cells(1, "K").Value = "Last Close Price"
        ws.Cells(1, "L").Value = "Yearly change"
        ws.Cells(1, "M").Value = "Percent Change"
        ws.Cells(1, "N").Value = "Total stock Volume"
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' Initialize the ticker count to zero
        TickerCount = 0
        ' Loop through each row of the data
        For i = 2 To lastRow
            ' Check if we have moved on to a new ticker
            If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
                ' If so, increment the ticker count and record the new ticker name
                TickerCount = TickerCount + 1
                Ticker = ws.Cells(i, "A").Value
                ' Record the open price for the new ticker
                OpenPrice = ws.Cells(i, "C").Value
                ' Record the last close price for the previous ticker
                If TickerCount > 1 Then
                    ws.Cells(TickerCount, "K").Value = LastClosePrice
    
                End If
                ' Reset the last close price for the new ticker
                LastClosePrice = ws.Cells(i, "F").Value
                ' Add a new row to the analysis sheet for the new ticker
                ws.Cells(TickerCount + 1, "I").Value = Ticker
                ws.Cells(TickerCount + 1, "J").Value = OpenPrice
            End If
            ' Calculate the percent change from open price for the current ticker
            LastClosePrice = ws.Cells(i, "F").Value
            ws.Cells(TickerCount + 1, "K").Value = LastClosePrice
            ws.Cells(TickerCount + 1, "L").Value = LastClosePrice - OpenPrice
            ws.Cells(TickerCount + 1, "M").Value = (LastClosePrice - OpenPrice) / OpenPrice
            ws.Cells(TickerCount + 1, "M").NumberFormat = "0.00%"
                        
            ' If this is the last row of the dataset, record the last close price for the current ticker
            If i = lastRow Then
                ws.Cells(TickerCount + 1, "M").Value = LastClosePrice
                            
            End If
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
    
                 Brand_Total = Brand_Total + ws.Cells(i, 7).Value
                 ws.Cells(TickerCount + 1, "N").Value = Brand_Total
                 ws.Cells(TickerCount + 1, "N").NumberFormat = "#,##0"
                 'Summary_Table_Row = Summary_Table_Row + 1
                 Brand_Total = 0
        
                Else
    
                Brand_Total = Brand_Total + ws.Cells(i, 7).Value
          
                 End If
            
            
        Next i
        ' Find ticker with highest gain
        Dim MaxGain As Double
        Dim MaxGainTicker As String
        MaxGain = 0
        For i = 2 To TickerCount
            If ws.Cells(i, "M").Value > MaxGain Then
                MaxGain = ws.Cells(i, "M").Value
                MaxGainTicker = ws.Cells(i, "I").Value
            End If
        Next i
    
        ' Add row for ticker with highest gain
        ws.Cells(6, "R").Value = MaxGainTicker
        ws.Cells(6, "Q").Value = "Max Gain"
        ws.Cells(6, "S").Value = MaxGain
        
        ' Find ticker with highest loss
        Dim MaxLoss As Double
        Dim MaxLossTicker As String
        MaxLoss = 0
        For i = 2 To TickerCount
            If ws.Cells(i, "M").Value < MaxLoss Then
                MaxLoss = ws.Cells(i, "M").Value
                MaxLossTicker = ws.Cells(i, "I").Value
            End If
        Next i
    
        ' Add row for ticker with highest gain
        
        ws.Cells(5, "R").Value = MaxLossTicker
        ws.Cells(5, "Q").Value = "Max Loss"
        ws.Cells(5, "S").Value = MaxLoss
        
        ' Find ticker with highest total volume
        Dim MaxVolume As Double
        Dim MaxVolumeTicker As String
        MaxVolume = 0
        For i = 2 To TickerCount
            If ws.Cells(i, "N").Value > MaxVolume Then
                MaxVolume = ws.Cells(i, "N").Value
                MaxVolumeTicker = ws.Cells(i, "I").Value
            End If
        Next i
        
        ' Add row for ticker with highest total volume
        ws.Cells(3, "R").Value = "Ticker"
        ws.Cells(4, "R").Value = MaxVolumeTicker
        ws.Cells(4, "Q").Value = "Greatest Total Volume"
        ws.Cells(4, "S").Value = MaxVolume
        
        ' Format "Yearly Change, M" column with red for negative and green for positive numbers
            Dim lastRow_ As Long
            lastRow_ = ws.Cells(Rows.Count, "L").End(xlUp).Row
    
        ' Format negative values as red
            With Range("L2:L" & lastRow_).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.ColorIndex = 3 ' Red color
            .Font.ColorIndex = 1 ' Black color
            End With
    
        ' Format positive values as green
            With Range("L2:L" & lastRow_).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
            .Interior.ColorIndex = 4 ' Green color
            .Font.ColorIndex = 1 ' Black color
        End With


    Next ws
    End Sub
    
    
    



