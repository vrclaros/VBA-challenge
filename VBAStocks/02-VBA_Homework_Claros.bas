Attribute VB_Name = "Module1"

Sub Main_Code()

Dim ws As Worksheet

    ' Loop through all sheets
    For Each ws In ActiveWorkbook.Worksheets

        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"

        'call Stock year to pull ticker and sum up volume
        
        Call Stock_Year
        
        'call Volume to sum up volume
        Call Volume
        
        'Call function to find max and min dates
        Call Find_dates
        
        'Call function to pull back price and format cells
        Call Price
        
        'call greatest function
        Call Greatest
        
    Next ws
    

End Sub



Public Sub Stock_Year()

'Define variables
Dim position As Integer
Dim max As Long
Dim open_price, close_price, vol As Double
Dim ticker, ticker_check As String

position = 2

max = Range("a1").End(xlDown).Row

'Loop through all ticker data and start summary in column I

For i = 2 To max

    ticker = Cells(i, 1).Value
    ticker_check = Cells(i + 1, 1).Value

    If ticker <> ticker_check Then
        Cells(position, 9) = ticker
        position = position + 1
    End If
    
Next i

End Sub

Public Sub Volume()

'define variables
Dim position As Integer
Dim max As Long
Dim open_price, close_price, vol As Double
Dim ticker, ticker_check As String

'calculate volume for the ticker names in column I
max_range = Range("i2").End(xlDown).Row
        
     
i = 2
For j = 2 To max_range
    ticker_check = Cells(j, 9).Value

      Do While ticker_check = Cells(i, 1).Value
        vol = vol + Cells(i, 7).Value
        i = i + 1
      Loop
       
    Cells(j, 12) = vol
    vol = 0
Next j
        
End Sub



Public Sub Find_dates()

'define variables
Dim year, month, day, position As Integer
Dim max, mmdd, mmdd_max, mmdd_min As Double
Dim open_price, close_price, vol As Double
Dim ticker, mmddString As String

max = Range("a1").End(xlDown).Row
max_range = Range("i2").End(xlDown).Row

'set min and max variables
mmdd_max = 0
mmdd_min = 1232

'loop through summary of ticker and find min and max dates
i = 2

For j = 2 To max_range
    ticker_check = Cells(j, 9).Value
        
        Do While ticker_check = Cells(i, 1).Value
        mmdd = CLng(Right(Cells(i, 2).Value, 4))
        
            If mmdd >= mmdd_max Then
                mmdd_max = mmdd
                mmddString = Cells(i, 2).Value
                Cells(j, 24) = mmddString
            End If
            
            If mmdd <= mmdd_min Then
                mmdd_min = mmdd
                mmddString = Cells(i, 2).Value
                Cells(j, 27) = mmddString
            End If

        i = i + 1
        Loop
        
    mmdd_max = 0
    mmdd_min = 1232
Next j
        

End Sub

Public Sub Price()

'define variables
Dim max, mmdd As Double
Dim price_open, price_close, yearly_change As Double
Dim ticker, ticker_check, mmddString As String
Dim mmdd_max, mmdd_min, mmdd_maxDate, mmdd_minDate As String

Dim perc_change As Double

max = Range("a1").End(xlDown).Row
max_range = Range("I2").End(xlDown).Row
lastRow = Range("k2").End(xlDown).Row

m = 2

For j = 2 To max_range
    ticker_check = Cells(j, 9).Value
    mmdd_min = Cells(j, 27).Value
      
      Do While ticker_check = Cells(m, 1).Value
        If mmdd_min = Cells(m, 2).Value Then
            price_open = Cells(m, 3).Value
            Cells(j, 28) = price_open
        End If
      
      m = m + 1
      Loop

On Error Resume Next
DoEvents
Next j

'set counter
x = 2

For n = 2 To max_range
    ticker_check = Cells(n, 9).Value
    mmdd_max = Cells(n, 24).Value

    Do While ticker_check = Cells(x, 1).Value
        If mmdd_max = Cells(x, 2).Value Then
            price_close = Cells(x, 6).Value
            Cells(n, 25) = price_close
        End If

      x = x + 1
    Loop
      
On Error Resume Next
DoEvents
Next n

'calculate
For y = 2 To max_range
    
    price_open = Cells(y, 28).Value
    price_close = Cells(y, 25).Value
    
    yearly_change = price_close - price_open
    Cells(y, 10) = yearly_change
    
    If price_open <> 0 Then
        perc_change = (yearly_change) / price_open
        Cells(y, 11) = perc_change
    ElseIf price_open = 0 Then
        Cells(y, 11).Value = 0
    End If

Next y

'formatting
Range("k2:k" & lastRow).NumberFormat = "0.00%"
Range("j2:j" & lastRow).NumberFormat = "0.00"


'conditional formatting

For i = 2 To max_range

    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value <= 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i


'clear values pulled
'Range("x1:x" & max_range).ClearContents
'Range("aa1:aa" & max_range).ClearContents
'Columns("X:AB").ClearContents



End Sub

Public Sub Greatest()

Dim max_range As Long
Dim max_change, min_change, max_vol, current_value As Double
Dim ticker_min, ticker_max, ticker_vol As String


max_range = Range("I2").End(xlDown).Row

max_change = 0
min_change = 0
max_vol = 0

For i = 2 To max_range

    current_value = Cells(i, 11).Value

    If current_value < min_change Then
        min_change = current_value
        ticker_min = Cells(i, 9).Value
    End If
    
    If current_value > max_change Then
        max_change = current_value
        ticker_max = Cells(i, 9).Value
    End If
Next i

For j = 2 To max_range
    current_value = Cells(j, 12).Value

    If current_value > max_vol Then
        max_vol = current_value
        ticker_vol = Cells(j, 9).Value
    End If

Next j

Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"

Cells(2, 15) = "Greatest % Increase"
Cells(2, 16) = ticker_max
Cells(2, 17) = max_change
Range("q2").NumberFormat = "0.00%"
Cells(3, 15) = "Greatest % Decrease"
Cells(3, 16) = ticker_min
Cells(3, 17) = min_change
Range("q3").NumberFormat = "0.00%"
Cells(4, 15) = "Greatest Total Volume"
Cells(4, 16) = ticker_vol
Cells(4, 17) = max_vol

End Sub


