Attribute VB_Name = "Module1"

Sub FormatingMacro()

    Dim ws As Worksheet
    
    Dim LastRow As Long
    LastRow = Range("J2").End(xlDown).Row

    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Gratest % Increase"
        ws.Range("O3").Value = "Gratest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
    LastRow = ws.Range("J2").End(xlDown).Row

    For i = 2 To LastRow
    
    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)

     End If
        
     If ws.Cells(i, 10).Value > 0 Then
     ws.Cells(i, 10).Interior.Color = RGB(102, 255, 102)

     End If
    
     Next i

     Next ws

        Worksheets("2018").Range("A:Q").Columns.AutoFit
        Worksheets("2019").Range("A:Q").Columns.AutoFit
        Worksheets("2020").Range("A:Q").Columns.AutoFit
    

End Sub




Sub TickerAnalysisFinal()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim LastClosePrice As Double
    Dim TickerCount As Long
    Dim i As Long
    Dim Brand_Total As Double
    Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Set ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
On Error Resume Next

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' Initialize the ticker count to zero
    TickerCount = 0
    ' Loop through each row of the data
    
    For i = 2 To LastRow
        ' Check if we have moved on to a new ticker
        If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
            ' If so, increment the ticker count and record the new ticker name
            TickerCount = TickerCount + 1
            Ticker = ws.Cells(i, "A").Value
            ' Record the open price for the new ticker
            OpenPrice = ws.Cells(i, "C").Value
            ' Reset the last close price for the new ticker
            LastClosePrice = ws.Cells(i, "F").Value
            ' Add a new row to the analysis sheet for the new ticker
            ws.Cells(TickerCount + 1, "I").Value = Ticker

        End If
        ' Calculate the percent change from open price for the current ticker
            LastClosePrice = ws.Cells(i, "F").Value
            ws.Cells(TickerCount + 1, "J").Value = LastClosePrice - OpenPrice
            ws.Cells(TickerCount + 1, "K").Value = (LastClosePrice - OpenPrice) / OpenPrice
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
            Volume = Volume + ws.Cells(i, 7).Value
            ws.Cells(TickerCount + 1, "L").Value = Volume
            Volume = 0
    
        Else

        Volume = Volume + ws.Cells(i, 7).Value
      
        End If
        
        
    Next i

On Error GoTo 0

Next ws

End Sub





Sub TickerSummary()
    
    Dim ws As Worksheet
    
    Dim TickerRange As Range
    Dim LastRowTicker As Long
    
    Dim VolumeRange As Range
    Dim LastRowVolume As Long
    
    Dim PercentchangeRange As Range
    Dim LastRowPercent As Long
    
    Dim TickerNameMax As String
    Dim TickerNameMin As String
    Dim TickerNameVolume As String
    
    Dim HighestVolume As Double
    Dim MaxPercent As Double
    Dim MinPercent As Double
    
    LastRowPercent = Range("K" & Rows.Count).End(xlUp).Row
    Set PercentchangeRange = Range("K2:K" & LastRowPercent)

    LastRowTicker = Range("I" & Rows.Count).End(xlUp).Row
    Set TickerRange = Range("I2:I" & LastRowTicker)
  
    LastRowVolume = Range("L" & Rows.Count).End(xlUp).Row
    Set VolumeRange = Range("L2:L" & LastRowVolume)
    
    MaxPercent = Application.WorksheetFunction.Max(PercentchangeRange)
    Range("Q2").Value = MaxPercent
    
    MinPercent = Application.WorksheetFunction.Min(PercentchangeRange)
    Range("Q3").Value = MinPercent
    
    HighestVolume = Application.WorksheetFunction.Max(VolumeRange)
    Range("Q4").Value = HighestVolume
    
    TickerNameMax = TickerRange(PercentchangeRange.Find(MaxPercent).Row - TickerRange.Row + 1, 1).Value
    Range("P2").Value = TickerNameMax
    
    TickerNameMin = TickerRange(PercentchangeRange.Find(MinPercent).Row - TickerRange.Row + 1, 1).Value
    Range("P3").Value = TickerNameMin

    TickerNameVolume = TickerRange(VolumeRange.Find(HighestVolume).Row - TickerRange.Row + 1, 1).Value
    Range("P4").Value = TickerNameVolume
    
    Range("Q2:Q3").Style = "Percent"



End Sub


Sub TickerSummary2()
    
    Dim ws As Worksheet
    Dim TickerRange As Range
    Dim LastRowTicker As Long
    Dim VolumeRange As Range
    Dim LastRowVolume As Long
    Dim PercentchangeRange As Range
    Dim LastRowPercent As Long
    Dim TickerNameMax As String
    Dim TickerNameMin As String
    Dim TickerNameVolume As String
    Dim HighestVolume As Double
    Dim MaxPercent As Double
    Dim MinPercent As Double
    
    For Each ws In ThisWorkbook.Worksheets
        
        On Error Resume Next 'skip any errors
        
            
            LastRowPercent = ws.Range("K" & ws.Rows.Count).End(xlUp).Row
            Set PercentchangeRange = ws.Range("K2:K" & LastRowPercent)
    
            LastRowTicker = ws.Range("I" & ws.Rows.Count).End(xlUp).Row
            Set TickerRange = ws.Range("I2:I" & LastRowTicker)
      
            LastRowVolume = ws.Range("L" & ws.Rows.Count).End(xlUp).Row
            Set VolumeRange = ws.Range("L2:L" & LastRowVolume)
            
            MaxPercent = Application.WorksheetFunction.Max(PercentchangeRange)
            ws.Range("Q2").Value = MaxPercent
            
            MinPercent = Application.WorksheetFunction.Min(PercentchangeRange)
            ws.Range("Q3").Value = MinPercent
            
            HighestVolume = Application.WorksheetFunction.Max(VolumeRange)
            ws.Range("Q4").Value = HighestVolume
            
            TickerNameMax = TickerRange(PercentchangeRange.Find(MaxPercent).Row - TickerRange.Row + 1, 1).Value
            ws.Range("P2").Value = TickerNameMax
            
            TickerNameMin = TickerRange(PercentchangeRange.Find(MinPercent).Row - TickerRange.Row + 1, 1).Value
            ws.Range("P3").Value = TickerNameMin
    
            TickerNameVolume = TickerRange(VolumeRange.Find(HighestVolume).Row - TickerRange.Row + 1, 1).Value
            ws.Range("P4").Value = TickerNameVolume
            
            ws.Range("Q2:Q3").Style = "Percent"
        
       
        
        On Error GoTo 0
        
    Next ws

End Sub


