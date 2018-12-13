Sub HomeWork()

'Declare Variables
    
    Dim LastRow1 As Long
    Dim LastRow2 As Long
    Dim Sh As Worksheet
    Dim I As Integer
    Dim j As Integer
    ' Declare variable for the first ticker
    Dim Ft As Long
    ' Declare variable for the last ticker
    Dim Lt As Long
    'Declare variable for first ticker Open value
    Dim Fv As Double
    ' Declare variable for the last ticker close value
    Dim Lv As Double
    ' declare variable to read values in column of unique tickers
    Dim t As String
   
' Variable j going to be counter of the sheets

    j = ThisWorkbook.Sheets.Count

' Name the columns for Ticker and Total Stock Volume
    For I = 1 To j
        ActiveWorkbook.Worksheets(I).Activate
        Set Sh = ActiveSheet
        
        ' Setting up first LastRow variable
        LastRow1 = Sh.Cells(Sh.Rows.Count, "A").End(xlUp).Row
    
        ' Naming Ticker and Total Stock Columns
        Sh.Range("I1").Value = "Ticker"
        Sh.Range("J1").Value = "Yearly Change"
        Sh.Range("K1").Value = "Percent Change"
        Sh.Range("L1").Value = "Total Stock Volume"
    
        'Coping the list of tickers from column A to Column I and removing duplicate
        Sh.Range("A2:A" & LastRow1).Copy
        Sh.Range("I2").PasteSpecial xlPasteValues
        Sh.Columns("I:I").RemoveDuplicates Columns:=1, Header:=xlNo
    
        ' Setting up second LastRow variable
        LastRow2 = Sh.Cells(Sh.Rows.Count, "I").End(xlUp).Row

            LastRow1 = Sh.Cells(Sh.Rows.Count, "A").End(xlUp).Row
        ' For statment to drill through each value in column A where
        
            For s = 2 To LastRow1
                If Sh.Cells(s, 9).Value <> "" Then
                t = Sh.Cells(s, 9).Value
                Ft = Sh.Range("A:A").Find(what:=t, after:=Sh.Range("A1"), lookat:=xlWhole).Row
                Lt = Sh.Range("A:A").Find(what:=t, after:=Sh.Range("A1"), searchdirection:=xlPrevious, lookat:=xlWhole).Row
                Fv = Sh.Cells(Ft, 3).Value
                Lv = Sh.Cells(Lt, 6).Value
                Sh.Cells(s, 10).Value = Lv - Fv
        ' Utilize error handler ince there is value of zero that raise error deviding with zero
                On Error Resume Next
                
                Sh.Cells(s, 11).Value = (Lv - Fv) / Fv
                
                End If
            Next
            
       ' Format Percent difference column
        Sh.Columns("K:K").Style = "Percent"
        Sh.Columns("K:K").NumberFormat = "0.00%"
        
       ' Create sum for Total Volume
        Sh.Range("L2").FormulaR1C1 = "=SUMIF(C[-11],RC[-3],C[-5])"

        ' Expanding the formulas from J to L and coping down to the last row of the column I.
        Sh.Range("L2").AutoFill Destination:=Range("L2:L" & LastRow2), Type:=xlFillDefault
        
       ' Format the columns width to expand so headers are fully visible
        Columns("I:L").EntireColumn.AutoFit
        
        ' Adding Conditional Formating to highlight negative values with red
        Range("J2:J1000000").Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        ' Adding Conditional Formating to highlight positive values with green
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
                .TintAndShade = 0
            End With
            
        Selection.FormatConditions(1).StopIfTrue = False
        'Adding the summary of greatest increase/decrease/total volume
        Sh.Range("O2").Value = "Greatest % Increase"
        Sh.Range("O3").Value = "Greatest % Decrease"
        Sh.Range("O4").Value = "Greatest Total Volume"
        Sh.Range("P1").Value = "Tikcer"
        Sh.Range("Q1").Value = "Value"
        Columns("O:O").EntireColumn.AutoFit
        Sh.Range("Q2").FormulaR1C1 = "=MAX(C[-6])"
        Sh.Range("Q3").FormulaR1C1 = "=MIN(C[-6])"
        Sh.Range("Q4").FormulaR1C1 = "=MAX(C[-5])"
		
        ' Used indexmatch unction to pull the ticker based on max/min and total volume
        Range("P2").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
        Range("P3").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
        Range("P4").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-4],0))"
        
        Columns("O:Q").EntireColumn.AutoFit
                           
    Next

End Sub


