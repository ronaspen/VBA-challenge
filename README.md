# VBA-challenge
Rohit Asopa VBA challenge
I had a lot of problems with saving the files after completing the coding. It kept saying that I needed to save the file to a Macro-safe folder. Yet I went into Trust settings and created Trust Locations, and still it didn’t allow me to save. 

Just in case the coding hasn’t come up in the file, here are the three Macros I used (I also attached screenshots):

Sub Stocks():
    
Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim Ticker_name As String
    Dim Vol_stock As Long
    Vol_stock = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim Open_Price As Double
    Open_Price = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    
    For i = 2 To LastRow
    
    Open_Price = Cells(2, 3).Value
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_name = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker_name
        Vol_stock = Cells(i, 7).Value
        Range("L" & Summary_Table_Row).Value = Vol_stock
        Yearly_Change = Cells(i, 6).Value - Open_Price
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        Percent_Change = (Yearly_Change / Open_Price) * 100
        Range("K" & Summary_Table_Row).Value = Percent_Change
        
        Summary_Table_Row = Summary_Table_Row + 1
        Open_Price = Cells(i + 1, 3).Value
            
        End If
        
        Next i
        
    Next ws
            
End Sub


Sub Color_Format():


Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
        
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
                
            ElseIf Cells(i, 10).Value = 0 Then
                Cells(i, 10).Interior.ColorIndex = 2
            
        
            
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            
            End If
         
         Next i
         Next ws
End Sub



Bonus:

Sub Greatest():

Dim ws As Worksheet
Dim highest_percent As Double
highest_percent = 0
Dim lowest_percent As Double
lowest_percent = 0
Dim highest_volume As Long
highest_volume = 0

    For Each ws In ThisWorkbook.Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

        If Cells(i + 1, 10).Value > Cells(i, 10).Value Then
        Cells(i + 1, 10).Value = highest_percent
        
        Cells(i + 1, 10).Value = Cells(2, 17)
        Cells(i + 1, 9).Value = Cells(2, 16)
        
        ElseIf Cells(i + 1, 10).Value < Cells(i, 10).Value Then
        Cells(i + 1, 10).Value = lowest_percent
        Cells(i + 1, 10).Value = Cells(3, 17)
        Cells(i + 1, 9).Value = Cells(3, 16)
        
    End If
    Next i
    
    
        
        
        
End Sub


Additionally, for a couple of the worksheets I got a bug saying ‘overflow’. I tried to change the value to CLng(), but it still said ‘overflow’.
I also tried to loop it for every worksheet in the code, but it didn’t work. So I just went to each Individual worksheet and ran the Macro.

