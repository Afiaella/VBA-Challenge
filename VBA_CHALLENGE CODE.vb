Sub Stock_Market ():

'Define the variables
Dim ticker as String 
ticker = " "
Dim yearly_open as Double 
yearly_open = 0
Dim yearly_close as Double 
yearly_close = 0
Dim total_vol as Double
total_vol = 0
Dim yearly_change as Double
yearly_change = 0
Dim percentage_change as Double
percentage_change = 0

Dim Summary_Table_Row as Long 
Summary_Table_Row =  2

Dim Lastrow As Long
Dim i As Long

Cells(1, 10).Value = "ticker" 
Cells(1, 11).Value = "yearly_change" 
Cells(1, 12).Value = "percentage_change" 
Cells(1, 13).Value = "total_vol"

'define last row
lastrow = Cells(Rows.Count, 1).End(xlUP).Row

 
'loop through all rows
For i = 2 To lastrow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
total_vol = total_vol + cells(i,7).value
 ticker = Cells(i,1).value
 yearly_open = Cells(2,3).value

 yearly_close = Cells(i,6).Value
 yearly_change = yearly_close - yearly_open
 percentage_change = (yearly_change / yearly_open) * 100

 Range("J" & Summary_Table_Row).value = ticker
 Range("K" & Summary_Table_Row).value = yearly_change
 Range("L" & Summary_Table_Row).value = (CStr(percentage_change) & "%")
 Range("M" & Summary_Table_Row).value = total_vol
  ' Add 1 to the summary table row count
Summary_Table_Row = Summary_Table_Row + 1
total_vol = 0

 Else
 total_vol = total_vol + Cells(i,7).Value

 End if 
 
 
Next i 



End Sub 




