Attribute VB_Name = "Module1"

' Ticker subrotine
Sub ticker_2018():


For Each ws In Worksheets
' declaring  a variable j as an index that stores the ticker
    Dim j As Integer
    j = 2
    ws.Cells(1, 9) = "Ticker"
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' looping through the column one
    For i = 2 To lastrow
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        
        Else
           j = j + 1
          ws.Cells(j, 9).Value = ws.Cells(i + 1, 1).Value
           
       End If
    Next i
 Next ws
      
End Sub

' Yearly change subroutine
Sub yearly_change():
    
    For Each ws In Worksheets
    ' declaring  a variable j as an index that stores the yearly changes
    ws.Cells(1, 10) = "Yearly Change"
    Dim m As Double ' index that keeps track of the ticker value at the beggining of the year
    Dim n As Double ' index that checks if the ticker is changed and increaments if its not changed
    Dim j As Double  ' index for each output
    
    m = 2
    n = 2
    j = 2
    
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
               
               n = n + 1
               
            Else
            
               ws.Cells(j, 10).Value = ws.Cells(n, 6).Value - ws.Cells(m, 3).Value
                   ' using case method to cheeck the yearly change is positive or negative
                    Select Case ws.Cells(j, 10).Value
                       Case Is > 0
                           ws.Cells(j, 10).Interior.ColorIndex = 4
                       Case Is < 0
                           ws.Cells(j, 10).Interior.ColorIndex = 3
                       Case Else
                           ws.Cells(j, 10).Interior.ColorIndex = 0
                    End Select
               m = n + 1
               n = n + 1
               j = j + 1
           End If
                
         Next i
    Next ws
End Sub

'yearly percentage change subroutine

Sub percentage_change():
   For Each ws In Worksheets
    ' declaring  a variable j that stores the ticker
    ws.Cells(1, 11) = "Percent Change"
    Dim m As Double ' index that keeps track of the ticker value at the beggining of the year
    Dim n As Double ' index that checks if the ticker is changed and increaments if its not changed
    Dim j As Double  ' index for each output
    
    m = 2
    n = 2
    j = 2
    
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
               
               n = n + 1
               
            Else
            
              
               ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / ws.Cells(m, 3).Value
               ws.Cells(j, 11).NumberFormat = "0.00%"
               m = n + 1
               n = n + 1
               j = j + 1
            End If
         Next i
    Next ws
End Sub

' Total stock volume subroutine

 Sub total_stock_volume():
   ' declaring  a variable j as an index that stores the total volume
   For Each ws In Worksheets
   ws.Cells(1, 12) = "Total Stock Volume"
   Dim j As Double
   Dim total_volume As Double
   j = 2
   ' setting the total volume to zero
   total_volume = 0
  
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
               
               total_volume = total_volume + ws.Cells(i, 7).Value  ' adding to the total_volume
               
            Else
               total_volume = total_volume + ws.Cells(i, 7).Value
               ws.Cells(j, 12).Value = total_volume
               total_volume = 0 ' resetting total volume to zero when the ticker name changes
               j = j + 1
            End If
         Next i
 Next ws
 
End Sub
' Greatest % Increase
Sub greatest_increase():
For Each ws In Worksheets
    Dim Highest As Double ' Declaring Highest  variable and setting it to the first sell value
    Highest = ws.Cells(2, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(2, 9).Value ' setting the name of the ticker to the first sell of the ticker name
   
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastrow
                ws.Cells(1, 16) = "Ticker"
                ws.Cells(1, 17) = "Value"
                ws.Cells(2, 15) = "Greatest % Increase"
                If ws.Cells(i, 11).Value > Highest Then
                    Highest = ws.Cells(i, 11).Value ' ressetting the value of the Highest
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value ' ressseting the name of the ticker as well
             
                End If
               
             Next i
              ws.Cells(2, 17).Value = Highest  ' storing the final value of Highest
              ws.Cells(2, 17).NumberFormat = "0.00%"
    Next ws
End Sub

' Greatest % Decrease
Sub greatest_Decrease():
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(3, 15) = "Greatest % Deacrease"
        Dim Lowest As Double
        Lowest = Cells(2, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(2, 9).Value
        
       For i = 2 To lastrow
           
           If ws.Cells(i, 11).Value < Lowest Then
               Lowest = ws.Cells(i, 11).Value
               ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
               
        
           End If
          
        Next i
         ws.Cells(3, 17).Value = Lowest
         ws.Cells(3, 17).NumberFormat = "0.00%"
     Next ws
End Sub


' Greatest % Decrease
Sub greatest_Total_Volume():
    For Each ws In Worksheets
    ws.Cells(4, 15) = "Greatest Total Volume"
    Dim Greatest_Total As Double
    Greatest_Total = ws.Cells(2, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(2, 9).Value
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastrow
                
                If ws.Cells(i, 12).Value > Greatest_Total Then
                    Greatest_Total = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
             
                End If
               
             Next i
              ws.Cells(4, 17).Value = Greatest_Total
   Next ws
End Sub



