Attribute VB_Name = "Module1"
Sub StockData()

'add sheet to end so macro can go to next tab after processing year 3
Sheets.Add After:=Sheets(Sheets.Count)

'Return to sheet one to begin the macro
Sheets(1).Select
Sheets(1).Activate

'Establish that there are 3 tabs that are to be processed
For j = 1 To 3

'Find last populated cell in column "A"
last = Cells(Rows.Count, 1).End(xlUp).Row

'Capture the tab name (year)
Dim year As String
year = ActiveSheet.Name

'Set Headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = year & " Volume" 'add year to the volume header description
    Cells(1, 12).Value = "Open"
    Cells(1, 13).Value = "Close"
    Cells(1, 14).Value = "Change"
    Cells(1, 15).Value = "%Change"

'Center Headers
    Range("J1:O1").HorizontalAlignment = xlCenter
    
'Set Volume column width to 15
Range("K:K").ColumnWidth = 15
    
'Set first Ticker, Volume, and open price
    Cells(2, 10).Value = Cells(2, 1)
    Cells(2, 11).Value = Cells(2, 7)
    Cells(2, 12).Value = Cells(2, 3)

'Declare Counter
    Dim counter As Integer
    counter = 2

'For Statement range (row three to the "last" populated cell in column "A"
    For i = 3 To last + 1

'Check if previous row ticker = current row ticker.
    If Cells(i - 1, 1) = Cells(i, 1) Then

'If true Update volume for current Ticker
    Cells(counter, 11).Value = Cells(i, 7) + Cells(counter, 11)
    
'If previous row <> current row. Then advance counter and add a new Ticker, update volumes for the new Ticker,
'update close price for the previous ticker, add price and percentage change (open to close) for the previous ticker,
'and update open price for the current ticker
    
    Else
        'advance counter
            counter = counter + 1
        'post new ticker
                Cells(counter, 10).Value = Cells(i, 1)
        'post new volume
                Cells(counter, 11).Value = Cells(i, 7) + Cells(counter, 11)
        'post close price of previous ticker
                Cells(counter - 1, 13).Value = Cells(i - 1, 6)
        'post price change from beginning to end of year
                Cells(counter - 1, 14).Value = Cells(counter - 1, 13) - Cells(counter - 1, 12)
        'post price change % from beginning to end of year
            'formula for % change between old price p1 and the new price p2 =
            '100 * (p2 - p1) / p1 (post blank if price open is 0)
                If Cells(counter - 1, 12) = 0 Then
                    Cells(counter - 1, 15).Value = ""
                Else
                    Cells(counter - 1, 15).Value = 100 * Cells(counter - 1, 14) / Cells(counter - 1, 12)
                End If
        'post open price on new ticker
                Cells(counter, 12).Value = Cells(i, 3)
        
End If

Next i

'formatting change and % change columns

'Find last populated cell in column "J"
last1 = Cells(Rows.Count, 10).End(xlUp).Row

'Define Cell formatting Range
Dim MyRange As Range
Set MyRange = Range("n2:O" & last1)

'Delete Existing Conditional Formatting from Range
MyRange.FormatConditions.Delete

'Apply Conditional Formatting Greater than or equal to zero
MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=0"
MyRange.FormatConditions(1).Interior.Color = RGB(0, 255, 0)


'Apply Conditional Formatting Less than zero
MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
MyRange.FormatConditions(2).Interior.Color = RGB(255, 0, 0)

'Find and then Post largest % gains and losses and largest volume
    'Declare variables
        Dim maxp As Double 'Max Percentage
        Dim minp As Double 'Min Percentage
        Dim maxv As Double 'Max Volume
        Dim tmaxp As String 'Ticker max percentage
        Dim tminp As String 'Ticker min percentage
        Dim tmaxv As String 'Ticker max volume
        
'set initial value
        maxp = Cells(2, 15)
        minp = Cells(2, 15)
        maxv = Cells(2, 11)
        tmaxp = Cells(2, 10)
        tminp = Cells(2, 10)
        tmaxv = Cells(2, 10)

'Enter Header and result descriptions
        Cells(1, 17).Value = "Desc"
        Cells(1, 18).Value = "Ticker"
        Cells(1, 19).Value = "Result"
        Cells(2, 17).Value = "Max % Change"
        Cells(3, 17).Value = "Min % Change"
        Cells(4, 17).Value = "Max Vol Change"

'Loop to evaluate changes and volumes
    For k = 2 To last1

'Evaluate for max percentage
If Cells(k, 15) > maxp Then
maxp = Cells(k, 15)
tmaxp = Cells(k, 10)
End If

'Evaluate for min percentage
If Cells(k, 15) < minp Then
minp = Cells(k, 15)
tminp = Cells(k, 10)
End If

'Evaluate for max volume
If Cells(k, 11) > maxv Then
maxv = Cells(k, 11)
tmaxv = Cells(k, 10)
End If

Next k

'Post min and max Tickers and values
Cells(2, 18).Value = tmaxp
Cells(3, 18).Value = tminp
Cells(4, 18).Value = tmaxv

Cells(2, 19).Value = maxp
Cells(3, 19).Value = minp
Cells(4, 19).Value = maxv

'Go to next sheet
ActiveSheet.Next.Select

'Pause to allow screen to update
Application.Wait Now + TimeSerial(0, 0, 5)
Next j

'Return to first tab at completion of code
Sheets(1).Select

End Sub




