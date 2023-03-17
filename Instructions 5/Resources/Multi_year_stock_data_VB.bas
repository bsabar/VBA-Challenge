Attribute VB_Name = "ModFirst"
Sub Stock_Market_Data()

'The Asks:
'I. Create a script that loops through all the stocks for one year and outputs the following information:
'(1) The ticker symbol
'(2) Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'(3) The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'(4) The total stock volume of the stock. The result should match the following image:

'II. Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'III. Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

'So first let's define all variables:
Dim ticker As String 'Define a variable for ticker
Dim date_open As Double 'Define a variable for date open
Dim date_close As Double 'Define avariable for date close
Dim Yearly_Change As Double 'Define a variable for yearly change
Dim Total_Stock_Volume As Double 'Define a variable for total stock volume
Dim Percent_Change As Double 'Define a variable for percent change
Dim begin As Integer 'Define a variable to set up a row to start
Dim ws As Worksheet 'Define variable of the worksheet to excute the code in all work sheet at once in the workbook

' Create a loop in all worksheet for the ask III:

For Each ws In Worksheets

    'Assign a column header for every calculations we are going to do:

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Assign intiger for the loop to start
    begin = 2
    first_i = 1
    Total_Stock_Volume = 0

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row     'Go to the last row of coumn A

    For i = 2 To EndRow  'For each ticker analyze and loop the yearly change, percent change, and total stock volume

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'If tickersymbol change or not equal to the previous one excute to record

                   ticker = ws.Cells(i, 1).Value      'Get the tickersymbol
                   first_i = first_i + 1    'Intiate the variable to go to the next ticker Alphabet
                   date_open = ws.Cells(first_i, 3).Value ' Get the value first day open form the column 3 or "C" and last day close of the year on column 6 or "F"
                  date_close = ws.Cells(i, 6).Value ' Get the value first day open form the column 3 or "C" and last day close of the year on column 6 or "F"

            For j = first_i To i       ' A for loop to sum the total stock volume using vol which is found in column 7 or "G"
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
            Next j

    
           If date_open = 0 Then    'When the loop get the value zero open the data
                Percent_Change = date_close
            Else
                Yearly_Change = date_close - date_open
                Percent_Change = Yearly_Change / date_open
            End If
    
    '*******************************************************
    
            'Get the values in the worksheet summary table
            ws.Cells(begin, 9).Value = ticker
            ws.Cells(begin, 10).Value = Yearly_Change
            ws.Cells(begin, 11).Value = Percent_Change

            'Use percentage format
            ws.Cells(begin, 11).NumberFormat = "0.00%"
            ws.Cells(begin, 12).Value = Total_Stock_Volume

            'In the data summery when the first row task completed go to the next row
            begin = begin + 1

            'Get back the variable to zero

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

            'Move i number to variable first_i
            first_i = i

        End If

    'Done the loop

    Next i

'The second summary table
     '*******************************************************

    'Go to the last row of column k
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Define variable to initiate the second summery table value
    Increase = 0
    Decrease = 0
    Greatest = 0

         For k = 3 To kEndRow    'find max/min for percentage change and the max volume Loop
            last_k = k - 1  'Define previous increment to check
            current_k = ws.Cells(k, 11).Value    'Define current row for percentage
            prevous_k = ws.Cells(last_k, 11).Value 'Define Previous row for percentage
            volume = ws.Cells(k, 12).Value 'greatest total volume row
            prevous_vol = ws.Cells(last_k, 12).Value       'Prevous greatest volume row

    '*******************************************************
 
            If Increase > current_k And Increase > prevous_k Then  'Find the increase
                Increase = Increase
                'define name for increase percentage
                'increase_name = ws.Cells(k, 9).Value
            ElseIf current_k > Increase And current_k > prevous_k Then
                Increase = current_k
                'define name for increase percentage
                increase_name = ws.Cells(k, 9).Value
            ElseIf prevous_k > Increase And prevous_k > current_k Then
                Increase = prevous_k
                'define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value
            End If

       '*******************************************************
          If Decrease < current_k And Decrease < prevous_k Then 'Find the decrease
                Decrease = Decrease     'Define decrease as decrease
         ElseIf current_k < Increase And current_k < prevous_k Then    'Define name for increase percentage
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
        ElseIf prevous_k < Increase And prevous_k < current_k Then
                Decrease = prevous_k
                decrease_name = ws.Cells(last_k, 9).Value
            End If

      '*******************************************************
      If Greatest > volume And Greatest > prevous_vol Then      'Find the greatest volume
        Greatest = Greatest

                'define name for greatest volume
                'greatest_name = ws.Cells(k, 9).Value

            ElseIf volume > Greatest And volume > prevous_vol Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value 'define name for greatest volume
            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                Greatest = prevous_vol
                greatest_name = ws.Cells(last_k, 9).Value     'define name for greatest volume
            End If
        Next k
    '*******************************************************
    ' Assign names for greatest increase,greatest decrease, and  greatest volume
    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "ticker Name"
    ws.Range("P1").Value = "Value"

    'Get for greatest increase, greatest increase, and  greatest volume ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest

    'Greatest increase and decrease in percentage format
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


    '*******************************************************
' Conditional formatting columns colors
   jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row 'The end row for column J
        For j = 2 To jEndRow
                If ws.Cells(j, 10) > 0 Then     'if greater than or less than zero
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

'Excute to next worksheet
Next ws
    '*******************************************************
End Sub

Sub Add_Formatting()
Attribute Add_Formatting.VB_Description = "Adding a few steps to practice recording Macros. Changing row 1 and column A to G with new color and font. "
Attribute Add_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
' Add_Formatting Macro
' Adding a few steps to practice recording Macros. Changing row 1 and column A to G with new color and font.
' I run this macro for 2019 and 2020 sheets too

    Range("A1:G1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Columns("A:G").Select
    Selection.Columns.AutoFit
    Columns("A:G").Select
End Sub
