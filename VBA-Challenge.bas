Attribute VB_Name = "Module1"
Sub VBA_Challenge()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Integer
    Dim yearlyChangeColumn As Integer
    Dim percentageChangeColumn As Integer
    Dim totalVolumeColumn As Integer
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim id As String
    Dim rng As Range
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Get the last row in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add Headers
        If ws.Cells(1, 9).Value <> "Ticker" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        End If
        
        ' Where to put data
        tickerColumn = 9
        yearlyChangeColumn = 10
        percentageChangeColumn = 11
        totalVolumeColumn = 12
        
        ' Remove duplicates and populate unique IDs in column I
        ws.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
        
        ' Loop through each row in the current worksheet
        For i = 2 To lastRow
            id = ws.Cells(i, 9).Value

            
            ' Update opening price if it's a new ID
            If ws.Cells(i, 2).Value <> "cells(i,1)" Then
                openingPrice = ws.Cells(i, 2).Value
            End If
            
            ' Closing price and total volume
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Output data for the last row of each ID
            If i = lastRow Or ws.Cells(i + 1, 9).Value <> id Then
                ' Yearly change
                yearlyChange = closingPrice - openingPrice
                ' Yearly change
                ws.Cells(i, yearlyChangeColumn).Value = yearlyChange
                ' Conditional formatting
                If yearlyChange < 0 Then
                    ws.Cells(i, yearlyChangeColumn).Interior.Color = RGB(255, 0, 0)
                ElseIf yearlyChange > 0 Then
                    ws.Cells(i, yearlyChangeColumn).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(i, yearlyChangeColumn).Interior.Color = RGB(255, 255, 0)
                End If
                
                ' Percentage change
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
                ' Percentage change
                ws.Cells(i, percentageChangeColumn).Value = percentageChange & "%"
                
                ' Total volume
                ws.Cells(i, totalVolumeColumn).Value = totalVolume
                
                totalVolume = 0
            End If
        Next i
    Next ws
    
    ' Formatting
    For Each ws In ThisWorkbook.Sheets
        ws.Columns("J").NumberFormat = "$#,##0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "#,##0"
    Next ws
    
    MsgBox "Project Complete."

End Sub

