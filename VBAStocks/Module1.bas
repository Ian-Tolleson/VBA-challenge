Attribute VB_Name = "Module1"
Sub Stock_Data()

    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        ws.Activate

    Rows("1:1").RowHeight = 30

    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("o2") = "Greatest % increase"
    Range("o3") = "Greatest % Decrease"
    Range("o4") = "Greatest total volume"
    
    Columns("O").ColumnWidth = 19



    Dim Ticker As String
    Dim PC As Double
    Dim TSV As Double
    Dim LastRow As Long
    Dim Storeticker As Boolean
    Dim FirstStart As Double
    Dim Difference As Double
    
    Storeticker = False
    FirstStart = Cells(2, 3).Value
    
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TSV = 0
    For i = 2 To LastRow
        
        'MsgBox (Cells(i, 7).Value)
        TSV = TSV + Cells(i, 7).Value
        
    If Storeticker <> True Then
    
        FirstStart = Cells(i, 3).Value
        Storeticker = True
            
        End If
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
        Difference = Cells(i, 6) - FirstStart
                
        Storeticker = False
        
        Ticker = Cells(i, 1).Value
        
        If Difference = 0 Or FirstStart = 0 Then
              PC = 0
        Else: PC = (Difference / FirstStart)
        
        End If
        
        
        Range("I" & Summary_Table_Row).Value = Ticker
        
        Range("L" & Summary_Table_Row).Value = TSV
                
        Range("J" & Summary_Table_Row).Value = Difference
        
        Range("K" & Summary_Table_Row).Value = PC
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        TSV = 0
        
        
    End If
    
    
    Next i
    
        Range("Q2") = WorksheetFunction.Max(Range("K2", "K" & LastRow))
        Range("Q3") = WorksheetFunction.Min(Range("K2", "K" & LastRow))
        Range("Q4") = WorksheetFunction.Max(Range("L2", "L" & LastRow))
           
        Range("P2") = WorksheetFunction.Index(Range("I2", "I" & LastRow), WorksheetFunction.Match(Range("Q2").Value, Range("K2", "K" & LastRow), 0))
        Range("P3") = WorksheetFunction.Index(Range("I2", "I" & LastRow), WorksheetFunction.Match(Range("Q3").Value, Range("K2", "K" & LastRow), 0))
        Range("P4") = WorksheetFunction.Index(Range("I2", "I" & LastRow), WorksheetFunction.Match(Range("Q4").Value, Range("L2", "L" & LastRow), 0))
    
    Next ws
    
    
End Sub

