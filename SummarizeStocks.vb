Option Explicit
Sub SummarizeStocks()
'=============================
'Define variables
'=============================
    Dim ticker As String
    Dim row As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim volume As LongLong
    Dim LastRow As Long
    Dim i As Long
    Dim j As Long
    Dim symbolcounter As Long
    
'=============================
'Determine opening Price for each ticker for the year
'=============================
    'count records in each ticker
    'set initial symbolcounter value
        symbolcounter = 0
    'Identify row with opening price by assigning a value of 1 to the symbol counter
    'NOTE: This apprpoach is slightly different than the master key approach.  However, I like this one better, because it allows for easier and quicker validation and troubleshooting
        LastRow = Cells(Rows.Count, 1).End(xlUp).row
            For j = 2 To LastRow
                If Cells(j + 1, 1).Value = Cells(j, 1).Value Then
                symbolcounter = symbolcounter + 1
                Cells(j, 8).Value = symbolcounter
                Else
                symbolcounter = 0
                End If
            Next j
           
'=============================
'Create column Headers in Summary Table
'=============================
'Note: As a Reporting and Analysis Best Practice, a program should never just return the Percent Change calculated value
'Best practice in accounting, finance and reporting is to include in the summary table the value upon which the final calculation is based
    Range("K1").Value = "Ticker"
    Range("L1").Value = "Open Price"
    Range("M1").Value = "Close Price"
    Range("N1").Value = "Yearly Change"
    Range("O1").Value = "Percent Change"
    Range("P1").Value = "Volume"
    
'=============================
'Create or Update Summary Table by Ticker
'=============================

' Keep track of the location for each ticker in the summary table
'Summary_Table_Row starts with 2 because we have column headers
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
'populate initial ticker
 'Set Initial Starting Point
    row = 2
    ticker = Cells(row, 1).Value
    volume = 0
' Loop through all ticker symbols
' Determine last row for the loop
    LastRow = Cells(Rows.Count, "A").End(xlUp).row
    For i = 2 To LastRow
    
    'calculate volume
       volume = volume + Cells(i, 7).Value
       
     'Deteremine open price for ticker using symbolcounter = 1
            If Cells(i, 8).Value = 1 Then
            openprice = Cells(i, 3)
            End If
            
        'Check if we are still within the same ticker symbol brand
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             
        'Ticker changing Set Ticker
            ticker = Cells(i, 1).Value
             'set close price
             closeprice = Cells(i, 6)
         
             
              'Print data to summary Table
                Range("K" & Summary_Table_Row).Value = ticker
                Range("L" & Summary_Table_Row).Value = openprice
                Range("M" & Summary_Table_Row).Value = closeprice
                Range("N" & Summary_Table_Row).Value = closeprice - openprice
                    If openprice <> 0 Then
                        Range("O" & Summary_Table_Row).Value = (closeprice - openprice) / openprice
                    Else
                        Range("O" & Summary_Table_Row).Value = 0
                    End If
                Range("P" & Summary_Table_Row).Value = volume
                Summary_Table_Row = Summary_Table_Row + 1
                volume = 0

            End If
      Next i
    
    '=============================
    'Conditional formatting for value greater than 0 or less than 0
    '=============================
       With Range("N2:N" & (Summary_Table_Row - 1))
        '=============================
        'conditional formatting for values greater than 0
        '=============================
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
   
        '=============================
        'conditional formatting for values less than 0
        '=============================
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False

 End With
'=============================
'Sheet clean up and table formatting
'=============================
    'Clear open price counter
        Columns("H:H").Select
        Selection.ClearContents
    'Set Decimal Places for percentage change
        Columns("O:O").Select
        Selection.NumberFormat = "0.00%"
    'Set format for Volume
        Columns("P:P").Select
        Selection.Style = "Comma"
        Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    'Clean up headers
        Range("K1:P1").Select
        Selection.Font.Bold = True
        Selection.Font.Underline = xlUnderlineStyleSingle
    'Save workbook
        ActiveWorkbook.Save
'=============================
'Message box that process is complete
'=============================
    MsgBox "Process Completed", vbInformation

End Sub



