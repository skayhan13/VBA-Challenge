Attribute VB_Name = "Module11"
Sub column_names()
'this subroutine will place column header names into newly specified column I
Dim column_names As String
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
End Sub

Sub ticker_list()
'this subroutine will find all unique ticker values in column A and copies it to column I
'declare your variables
    Dim open_price As Double
    Dim closing_price As Double
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim i As Long
    Dim j As Long 'this will be the index for the summary tables
    Dim last_row As Long
    j = 1


    last_row = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To last_row
        ticker = Cells(i, 1).Value
            If ticker <> Cells(i + 1, 1).Value Then
                Cells(i, 9).Value = ticker
                j = j + 1
                Cells(j, 9).Value = Cells(i, 1)
                Cells(j + 1, 9).Value = Cells(i + 1, 1)
                'Exit For
            End If
        Next i

End Sub

Sub total_stock_volume()
'this subroutine will sum the volume for each stock
    Dim i As Double 'for loop variable for ticker row
    Dim j As Double 'for loop variable for unique row
    Dim last_row As Double
    Dim total_volume As Double



    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    total_volume = 0
    j = 2
    Cells(j, 9).Value = Cells(2, 1).Value
    
    For i = 2 To last_row
    
        If Cells(i, 1).Value = Cells(j, 9).Value Then
            total_volume = total_volume + Cells(i, 7)
        Else
            Cells(j, 12).Value = total_volume
            total_volume = 0 + Cells(i, 7).Value
            j = j + 1
            Cells(j, 9).Value = Cells(i, 1).Value
        End If
    Next i
End Sub

Sub analysis()
'this sub routine will find yearly change and total percent change for each unique ticker
Dim i As Double
Dim last_row As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
'Dim PercentChange As LongLong
Dim YearlyChange As Double
Dim top_row As Double


last_row = Cells(Rows.Count, 1).End(xlUp).Row
top_row = 2

    For i = 2 To last_row
   If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
     
      ClosePrice = Cells(i, 6)
      OpenPrice = Cells(i, 3)
      YearlyChange = ClosePrice - OpenPrice
      PercentChange = Format((ClosePrice - OpenPrice) / (OpenPrice + 0.0000001), "0.##%") 'this will address any open prices of 0 and format it to 2 decimal places
      
      Cells(top_row, 11).Value = PercentChange
      Cells(top_row, 10).Value = YearlyChange
      top_row = top_row + 1 'this ensures that the values found will be properly placed in correct columns
      
    End If
       
    Next i 'tells the machiene to keep looking for all values

End Sub

Sub percent_change()
'this subroutine uses interior color formatting for yearly change less than 0 and greater than 0

Dim last_unique_row As Double
Dim j As Double

last_unique_row = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To last_unique_row
        If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
        End If
   Next j
End Sub

Sub combinedata()
    'this subroutine will combine all subroutines to allow to run across multiple worksheets

WS_Count = ActiveWorkbook.Worksheets.Count

For i = 1 To WS_Count
    Sheets(i).Activate

    Call column_names

    Call ticker_list

    Call total_stock_volume

    Call analysis

    Call percent_change

Next i

End Sub
