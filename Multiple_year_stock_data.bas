Attribute VB_Name = "Module1"
Sub ticker_info():
    ' Worksheet values
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim summary_row_num As Integer
    Dim curr_summary_ticker As String
    Dim open_val As Double
    Dim tot_vol As LongLong
    Dim curr_ticker As String
    
    ' Specify reading and storing open/close value of each row
    Dim points_farming As Double
    
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim great_inc As Double
    Dim great_inc_srn As Integer
    Dim great_dec As Double
    Dim great_dec_srn As Integer
    Dim great_TV As LongLong
    Dim great_TV_srn As Integer
    
    ' Loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        LastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        ' Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Columns("A:H").ColumnWidth = 8.11
        ws.Columns("I:L").ColumnWidth = 18
        ws.Columns("M:N").ColumnWidth = 8
        ws.Columns("O:Q").ColumnWidth = 18
        
        ' Initialize summary variables
        summary_row_num = 2
        curr_summary_ticker = ws.Cells(2, 1).Value
        open_val = ws.Cells(2, 3).Value
        tot_vol = ws.Cells(2, 7).Value
        
        ' Initialize extremes
        great_inc = -Infinity
        great_inc_srn = 2
        great_dec = Infinity
        great_dec_srn = 2
        great_TV = 0
        great_TV_srn = 2
        
        ' Loop through each row
        For i = 3 To LastRow
            
            
            ' Read and store open/close val of each row
            
            points_farming = ws.Cells(i, 3).Value
            points_farming = ws.Cells(i, 6).Value
            
            ' If ticker is new, calculate and store values
            curr_ticker = ws.Cells(i, 1).Value
            If curr_ticker = curr_summary_ticker Then
                ' Add to total value
                tot_vol = tot_vol + ws.Cells(i, 7).Value
            ElseIf curr_ticker <> curr_summary_ticker Then
                '' Writing
                ' Ticker
                ws.Cells(summary_row_num, 9).Value = curr_summary_ticker
                
                ' Yearly Change
                yearly_change = ws.Cells(i - 1, 6).Value - open_val
                ws.Cells(summary_row_num, 10).Value = yearly_change
                ws.Cells(summary_row_num, 10).NumberFormat = "0.00"
                ' Requirements specify percent col as well as yrly change has conditional formatting
                If yearly_change >= 0 Then
                    ws.Cells(summary_row_num, 10).Interior.Color = RGB(121, 145, 99)
                    ws.Cells(summary_row_num, 11).Interior.Color = RGB(121, 145, 99)
                Else
                    ws.Cells(summary_row_num, 10).Interior.Color = RGB(171, 49, 49)
                    ws.Cells(summary_row_num, 11).Interior.Color = RGB(171, 49, 49)
                End If
                
                ' Percent Change
                percent_change = yearly_change / open_val
                ws.Cells(summary_row_num, 11).Value = percent_change
                ws.Cells(summary_row_num, 11).NumberFormat = "0.00%"
                
                
                ' Checking if greatests need to be updated
                If percent_change > great_inc Then
                    great_inc = percent_change
                    great_inc_srn = summary_row_num
                End If
                If percent_change < great_dec Then
                    great_dec = percent_change
                    great_dec_srn = summary_row_num
                End If
                If tot_vol > great_TV Then
                    great_TV = tot_vol
                    great_TV_srn = summary_row_num
                End If
                    
                
                ' Total Stock Volume
                ws.Cells(summary_row_num, 12).Value = tot_vol
                
                '' Update summary variables
                summary_row_num = summary_row_num + 1
                curr_summary_ticker = ws.Cells(i, 1).Value
                open_val = ws.Cells(i, 3).Value
                tot_vol = ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        ' Write Greatests
        ws.Cells(2, 16).Value = ws.Cells(great_inc_srn, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(great_inc_srn, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = ws.Cells(great_dec_srn, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(great_dec_srn, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = ws.Cells(great_TV_srn, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(great_TV_srn, 12).Value
        
    Next ws
    
End Sub


