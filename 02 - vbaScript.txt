Sub ticker()
    ' First loop to go through sheets
    For Each sheet In Worksheets
        ' Declare variables
        Dim stockSingle As String
        Dim stockGroup As String
        Dim stockSingleCount As Double
        Dim stockGroupCount As Double
        Dim volTotal As Double
        
        sheet.Activate
        sheet.Range("I1:J3200").Clear
        
        ' Create a list of Ticker Names and Total Stock Volume
        ActiveSheet.Range("A2:A800000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("i2"), Unique:=True
        
        'My run time was six hours at 800000.
        stockSingleCount = 800000
        stockGroupCount = 3200
        'stockSingleCount = sheet.Range("A800000").End(xlUp).Row
        'stockGroupCount = sheet.Range("I3000").End(xlUp).Row
        
        ' First loop through Ticker Names
        For i = 2 To stockGroupCount
            stockGroup = Cells(i, 9).Value
            volTotal = Cells(i, 10).Value
            
            ' Second loop through <ticker>
            For j = 2 To stockSingleCount
                stockSingle = Cells(j, 1).Value
                
                ' Add volume for that day to the main total
                If stockSingle = stockGroup Then
                    volTotal = volTotal + Cells(j, 7).Value
                
                End If
            
            Next j
            
            ' Show <vol> total
            Cells(i, 10).Value = volTotal
        
        Next i
        
        ' Add column names
        Range("I2").Value = "Ticker Names"
        Range("J2").Value = "Total Stock Volume"
        
    Next

End Sub
