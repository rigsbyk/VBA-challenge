Attribute VB_Name = "Module11"

'   Create sub routine
Sub mkt():

'   Declaring and setting all variables for workbook
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

'   Creating a for each loop to go through each worksheet in the workbook and activate and perform the code automatically
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'   Declare all variables
    Dim counter As Double
    Dim lastrow As Long
    Dim lastrowtv As Long
    Dim lastrowv As Long
    Dim percentchange As Double
    Dim total_vol As Double
    Dim yropen As Double
    Dim yrend As Double
    Dim yrchange As Double
   
   '    setting the counter and yrend initially at zero
    counter = 0
    yrend = 0
    
    '   sets the headers for Columns "I" through "L"
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
      
    '    this allows the code to find and execute through the last row of data without having to hard code what the last row is. It makes it more dynamic
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
    '   need a loop to go through and each row; Since row 1 is headers, then the data would start on row 2
    For i = 2 To lastrow
            
            '  Setting Column "C" as the data for yropen
             If yropen = 0 Then
                yropen = Cells(i, 3).Value
            End If
            
             '  Comparing the current cell to the next cell. If different, do the below, the same it goes in the else statement
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                 '  Total Stock volume
                total_vol = total_vol + Cells(i, 7)
           
                '   Stating what data to use for yrend
                yrend = Cells(i, 6).Value
                
                'Giving it a formula to calculate for Yearly change: Yearly change = Close - Open
                yrchange = yrend - yropen
                
                '   Place the ticker values in Cells(i,9), which is Column "I"
                Cells(counter + 2, 9).Value = Cells(i, 1).Value
                
                '   Place the yearly change values in Cells(i, 10), which is Column "J"
                Cells(counter + 2, 10).Value = yrchange
                
                '   Placing a conditional on the percent formula so that if yropen does = 0 it sets the percentchange to 0 instead of erroring out
                    If yropen = 0 Then
                        percentchange = 0
                    Else
                        '   Find the percentchange formula
                        percentchange = (yrchange / yropen)
                        
                    End If
                '   Place the percent change values in cells(i, 11), which is Column "K"
                Cells(counter + 2, 11).Value = percentchange
                Cells(counter + 2, 11).NumberFormat = "0.00%"
                
                '   Place the total volume values in Cells(i,12), Which is Column "L"
                Cells(counter + 2, 12).Value = total_vol
                
                '   Placing a condition if the yrchange is greater than 0, aka positive set to green. If it is less than 0, aka negative then set to red
                 If yrchange > 0 Then
                    Cells(counter + 2, 10).Interior.ColorIndex = 4
                Else
                    Cells(counter + 2, 10).Interior.ColorIndex = 3
                End If
             
                '  Stating how the counter will increment
                counter = counter + 1
                
                '   This sets the total_vol back at 0 before the ticker changes
                total_vol = 0
                '   This sets the yropen back at 0 before the ticker changes
                yropen = 0
                 

            Else
                '   If the ticker remains the same just keep totaling
                 total_vol = total_vol + Cells(i, 7)
                 
            '   Ends the if statement
            End If
                        
        '   Ends the  loop i
        Next i
        
        '   Find greatest % increase and greatest % decrease headers
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrase"
        Cells(4, 15).Value = "Greatest Total Volume'"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
       '    Set another lastrowv
        lastrowv = Cells(Rows.Count, 11).End(xlUp).Row
       
        '   Create a new loop to find greatest % increase and greatest % decrease
        For j = 2 To lastrowv
        
            '   Look through Column "K" and find the max value and then...
            If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrowv)) Then
                '   Place the associated ticker in cell(2, 16)
                Cells(2, 16).Value = Cells(j, 9).Value
                '   Place the max percent value in cell(2, 17)
                Cells(2, 17).Value = Cells(j, 11).Value
                '    Sets the value to percent
                Cells(2, 17).NumberFormat = "0.00%"
            End If
            
            '   Look through Column "K" and find the min value and then...
            If Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrowv)) Then
                '   Place the associated ticker in cell(3, 16)
                Cells(3, 16).Value = Cells(j, 9).Value
                '   Place the min value in cell(3, 17)
                Cells(3, 17).Value = Cells(j, 11).Value
                '   Sets the value to percent
                 Cells(3, 17).NumberFormat = "0.00%"
            End If
            
        '   Ends for loop j
        Next j
        
        '   Set another last row for the max volume
         lastrowtv = Cells(Rows.Count, 12).End(xlUp).Row
        
        '   Start new loop to find max volume
        For k = 2 To lastrowtv
        
            '   Look through column "L" and find the max volume and then...
             If Cells(k, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrowtv)) Then
                '   Place the associated ticker in cell(4, 16)
                Cells(4, 16).Value = Cells(k, 9).Value
                '   Place the max volume value in cell(4, 17)
                Cells(4, 17).Value = Cells(k, 12).Value
                
            End If
            
        '   End loop k
        Next k
        
     '  This sets cell A1 of each sheet to "1"
    ws.Cells(1, 1) = "<ticker>"
    
'   Ends for each loop
Next

'   Activate the worksheet that was originally active
starting_ws.Activate

'   Ends the sub routine
End Sub


