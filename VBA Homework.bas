Attribute VB_Name = "Module1"
Sub Ticker()

Dim TickerName As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStock As Double
Dim SummaryTableRow As Integer
Dim LastRow As Long
Dim InitialOpenVariable As Double
Dim ClosingVariable As Double
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
SummaryTableRow = 2
TotalStock = 0
InitialOpenVariable = Cells(2, 3).Value


    For i = 2 To LastRow
    
    'Set TotalStock
    TotalStock = TotalStock + Cells(i, 7).Value
            
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Print TotalStock
            Cells(Str(SummaryTableRow), 14).Value = TotalStock
            
            'Set and print Ticker Name
            TickerName = Cells(i, 1).Value
            Cells(Str(SummaryTableRow), 9).Value = TickerName
            
            'Set and print Open Variable and Closing Variable
            Cells(Str(SummaryTableRow), 10).Value = InitialOpenVariable
            ClosingVariable = Cells(i, 6).Value
            Cells(Str(SummaryTableRow), 11).Value = ClosingVariable
    
            'Set and print Yearly Change
            YearlyChange = Cells(i, 6).Value - InitialOpenVariable
            Cells(Str(SummaryTableRow), 12).Value = YearlyChange
            
            'Set and print PercentChange
            If InitialOpenVariable <> 0 And Cells(i, 6).Value <> 0 Then
                PercentChange = YearlyChange / InitialOpenVariable
                Cells(Str(SummaryTableRow), 13).Value = PercentChange
                Cells(Str(SummaryTableRow), 13) = Format(PercentChange, "0.00%")
            End If
            
        InitialOpenVariable = Cells(i + 1, 3).Value
        TotalStock = 0
        
            If YearlyChange < 0 Then
            Cells(Str(SummaryTableRow), 12).Interior.ColorIndex = 3
            Else
            Cells(Str(SummaryTableRow), 12).Interior.ColorIndex = 4
            End If
    
        SummaryTableRow = SummaryTableRow + 1
    
        End If
        
           
    Next i
    
Next ws

End Sub

