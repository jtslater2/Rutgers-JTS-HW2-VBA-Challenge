Attribute VB_Name = "Module1"
Sub Stockmarket():


'Variables

Dim i, j, k As Long

Dim Openingprice, Closingprice, Yearlychange, PercentChange As Double

Dim Stockvolumn As Double

Dim tickersymb As String

Dim stackloc As Double

Dim LastRow As Double

Dim ws As Worksheet

Dim wscount As Integer

wscount = Worksheets.Count

    'Test MsgBox ("worksheet count = " & wscount)

For k = 1 To wscount

Worksheets(k).Activate

i = 0

    'Test MsgBox ("i value  this should be zero   " & i)


'Setup of the Worksheets

    'Set Price cells to Numbers
    Columns("C:F").Select
    Selection.NumberFormat = "0.00"
           
    'Set Column K Q2 & Q3 to Percentage
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    Range("H1").Select
         
    'Add Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Adjust Column Widths
    Columns("L:L").ColumnWidth = 20
    Columns("O:O").ColumnWidth = 22
    Columns("G:G").ColumnWidth = 16
    Columns("Q:Q").ColumnWidth = 16
    Columns("J:J").ColumnWidth = 13
    Columns("K:K").ColumnWidth = 16
  
'Set size of i loop by getting lastrow value
LastRow = ActiveSheet.UsedRange.Rows.Count
    'Test MsgBox ("lastrow " & LastRow)

'Take first ticker symbol and put it in "I2"
    Cells(2, 9).Value = Cells(2, 1).Value
    'Test MsgBox (Range("i2").Value)

'Set the stack location counter to create Ticker chart
    stackloc = 2

'Take first ticker opening value & volume and put it in Openingprice & Stockvolume
    Openingprice = Cells(2, 3).Value
    Stockvolume = Cells(2, 7).Value
    
    'Test MsgBox ("OpeningPrice " & Openingprice)
    'Test MsgBox ("Stockvolume " & stockvolume)
    
For i = 2 To LastRow - 1
        
'Ticker value matches Ticker value one line down
    If Cells(i, 1).Value = Cells(i + 1, 1) Then
        
      'Adding matching line below to Stock Volume
      Stockvolume = Stockvolume + Cells(i + 1, 7).Value
      Cells(stackloc, 12).Value = Stockvolume
        
    End If
        
'Ticker value does not match Ticker value one line down
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            'Get closing price
            Closingprice = Cells(i, 6).Value
        
            'Calculate Yearly change & print
            Yearlychange = Closingprice - Openingprice
            Cells(stackloc, 10).Value = Yearlychange
        
            'Calculate Percent change & print
            If Openingprice <> 0 Then
            PercentChange = Yearlychange / Openingprice
            Cells(stackloc, 11).Value = PercentChange
            Else
            Cells(stackloc, 11).Value = 0
            End If
        
            'Set Next Ticker, Opening Price & Stock Volume
            tickersymb = Cells(i + 1, 1).Value
                'Test MsgBox ("tickersymb " & tickersymb)
            
            'Increment stackloc counter
            stackloc = stackloc + 1
        
            'Print tickersymbol in chart
            Cells(stackloc, 9).Value = tickersymb
        
                'Test MsgBox ("correct ticker in correct place??")
        
        
            Openingprice = Cells(i + 1, 3).Value
        
            Stockvolume = Cells(i + 1, 7).Value
        
                'Test MsgBox ("stackloc" & stackloc)
        
        End If
        
    Next i
            'Test MsgBox ("i value after last next i " & i)
            'Test MsgBox ("lastrow & I should match " & LastRow)
            
        If i = LastRow Then
        
            'Add last Stock Volume
            Stockvolume = Stockvolume + Cells(i, 7).Value
                
            'Fill in the last line for Closing price
            Closingprice = Cells(i, 6).Value
        
            'Calculate Yearly change & print
            Yearlychange = Closingprice - Openingprice
            Cells(stackloc, 10).Value = Yearlychange
        
            'Calculate Percent change & print
            If Openingprice <> 0 Then
            PercentChange = Yearlychange / Openingprice
            Cells(stackloc, 11).Value = PercentChange
            Else
            Cells(stackloc, 11).Value = 0
            End If
        
        End If
        

'Add Ticker & values for Greatest Increase/Decrease & Greatest Volume

For i = 2 To stackloc

  'Look for the Greatest Increase - Print Ticker & Value
    If Cells(i, 11).Value > Range("Q2").Value Then
    Range("Q2").Value = Cells(i, 11)
    Range("P2").Value = Cells(i, 9)
    End If
    
   'Look for the Greatest Decrease - Print Ticker & Value
    If Cells(i, 11).Value < Range("Q3").Value Then
    Range("Q3").Value = Cells(i, 11)
    Range("P3").Value = Cells(i, 9)
    End If
   
   'Look for Greatest Volume - Print Ticker & Value
    If Cells(i, 12).Value > Range("Q4").Value Then
    Range("Q4") = Cells(i, 12).Value
    Range("P4").Value = Cells(i, 9)
    End If
    
Next i

'Conditional Format Cells using for each cell in range

Dim stackloc_j As String
stackloc_j = "J2:" & "J" & stackloc

For Each Cell In Range(stackloc_j)
    
    'Color Cell Red for (-)
    If Cell.Value < 0 Then
    Cell.Interior.ColorIndex = 3
    End If
    
    'Color Cell Green for (+)
    If Cell.Value > 0 Then
    Cell.Interior.ColorIndex = 4
    End If
        
Next Cell
    
Next k

End Sub



