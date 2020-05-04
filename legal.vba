Sub Button7_Click()


'''MAKE SURE ONLY ONE INSTANCES OF ACCOUNT NUMBER EXISTS
'''MAKE SURE ONLY ONE INSTANCE OF EACH CATEGORY LISTED EXISTS. THIS IS WHAT THE POINT OF THE "FLOYD'S SHEET" TAB IS DOING
'''IT IS SUMMARIZING THE DATA TO MAKE SURE THERE IS NOT TWO VERY SIMILAR ITEMS THAT DO NOT CONSOLIDATE
'''EXAMPLE  "PAYING" VS " PAYING" WHERE THE SECOND ITEM HAS A BEGINNING SPACE



With ActiveSheet
' sets the last row variable to the last row in the first column
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

errorCheck = 0
' loop over all cells '
For i = 2 To (lastRow)
   
   ' trim spaces on either side of
   Cells(i, "D").Value = Trim(Cells(i, "D").Value)


   ' check row "E" for duplicate account numbers
   currentValue = Cells(i, 5)
   If (currentValue > 1) Then
   
        MsgBox ("duplicate account occured for account number: " & Cells(i, 1) & " in row " & i)
        errorCheck = 1
   
        
    End If
Next i

If (errorCheck = 1) Then
    MsgBox ("Data errors")
    Else: MsgBox ("Data OK")
    End If
End With


End Sub




Sub Button117_Click()


' MAKE SURE NO DATA EXISTS PAST THE LAST ROW OF DATA PASTED
' MAKE SURE FORMULAS WORKING CORRECT

With ActiveSheet
' sets the last row variable to the last row in the first column
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' counts that there is the correct number of rows in the seet
If (lastRow <> 720) Then
    MsgBox ("The last row is" & lastRow & "which should be row 720")
    Else: MsgBox ("number of Rows OK")
End If

errorCount = 0

' checks if there are any errors in the sheet
Dim sheetRange As Range
Set sheetRange = Range("A5:DL720")
    
    Dim cell As Range
    
    For Each cell In sheetRange
        errorVal = IsError(cell.Value)
        If (errorVal = True) Then
            errorCount = errorCount + 1
            MsgBox ("Error in cell: " & cell.Address)
         End If

        
        
    Next cell
MsgBox ("Error Count is: " & errorCount)
End With
End Sub


Sub datadroptable_Button1_Click()


  On Error Resume Next
  
  Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A4:DL" & Lastrow).Cells.SpecialCells(xlCellTypeConstants).ClearContents

  'Cells.SpecialCells(xlCellTypeConstants).ClearContents

End Sub


Sub datadroptable_Button1_Click()
With ActiveSheet
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Range("A4:DL" & Lastrow).Cells.SpecialCells(xlCellTypeConstants).ClearContents
'Range("A4:DL" & Lastrow).Clear
End With
End Sub

Sub Button9_Click()
  ActiveWorkbook.RefreshAll
End Sub


Sub Button7_Click()


'''MAKE SURE ONLY ONE INSTANCES OF ACCOUNT NUMBER EXISTS
'''MAKE SURE ONLY ONE INSTANCE OF EACH CATEGORY LISTED EXISTS. THIS IS WHAT THE POINT OF THE "FLOYD'S SHEET" TAB IS DOING
'''IT IS SUMMARIZING THE DATA TO MAKE SURE THERE IS NOT TWO VERY SIMILAR ITEMS THAT DO NOT CONSOLIDATE
'''EXAMPLE  "PAYING" VS " PAYING" WHERE THE SECOND ITEM HAS A BEGINNING SPACE



With ActiveSheet
' sets the last row variable to the last row in the first column
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

errorCheck = 0
' loop over all cells '
For i = 2 To (Lastrow)
   
   ' trim spaces on either side of
   Cells(i, "D").Value = Trim(Cells(i, "D").Value)


   ' check row "E" for duplicate account numbers
   currentValue = Cells(i, 5)
   If (currentValue > 1) Then
   
        MsgBox ("duplicate account occured for account number: " & Cells(i, 1) & " in row " & i)
        errorCheck = 1
   
        
    End If
Next i

If (errorCheck = 1) Then
    MsgBox ("Data errors")
    Else: MsgBox ("Data OK")
    End If
End With


End Sub




Sub Button117_Click()


' MAKE SURE NO DATA EXISTS PAST THE LAST ROW OF DATA PASTED
' MAKE SURE FORMULAS WORKING CORRECT

With ActiveSheet
' sets the last row variable to the last row in the first column
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' counts that there is the correct number of rows in the seet
If (Lastrow <> 720) Then
    MsgBox ("The last row is" & Lastrow & "which should be row 720")
    Else: MsgBox ("number of Rows OK")
End If

errorCount = 0

' checks if there are any errors in the sheet
Dim sheetRange As Range
Set sheetRange = Range("A5:DL720")
    
    Dim cell As Range
    
    For Each cell In sheetRange
        errorVal = IsError(cell.Value)
        If (errorVal = True) Then
            errorCount = errorCount + 1
            MsgBox ("Error in cell: " & cell.Address)
         End If

        
        
    Next cell
MsgBox ("Error Count is: " & errorCount)
End With
End Sub

Sub Button3_Click()
'Inserting a Column at Column B
With ActiveSheet
'Range("E1").EntireColumn.Insert
'Range("E4:E9").Style = "Currency"
'Range("E19:E24").Style = "Currency"
'Range("Q1").EntireColumn.Delete

End With
End Sub
Sub Button4_Click()

With ActiveSheet
'Delete a column at Column B
'Range("E1").EntireColumn.Delete
End With
End Sub
Sub Button1_Click()
'Refresh pivot table data
ActiveWorkbook.RefreshAll
End Sub
Sub Button5_Click()
'get values from CFS AGING SUMMARY


' copy all data from rows E to O to F to P
Worksheets("AGING TRACKING").Range("E3:O24").Copy _
    Destination:=Worksheets("AGING TRACKING").Range("F3:P24")

Worksheets("CFS AGING SUMMARY").Range("J5:J10").Copy Worksheets("AGING TRACKING").Range("E4:E9")
Worksheets("CFS AGING SUMMARY").Range("I5:I10").Copy Worksheets("AGING TRACKING").Range("E19:E24")
With ActiveSheet
' create new percentages
'Cells("E", 11).NumberFormat = "0.00%"

ActiveSheet.Range("E11:E16").NumberFormat = "0.00%"



Cells(11, "E").Value = (Cells(4, "E").Value) / (Cells(9, "E").Value)
Cells(12, "E").Value = (Cells(5, "E").Value) / (Cells(9, "E").Value)
Cells(13, "E").Value = (Cells(6, "E").Value) / (Cells(9, "E").Value)
Cells(14, "E").Value = (Cells(7, "E").Value) / (Cells(9, "E").Value)
Cells(15, "E").Value = (Cells(8, "E").Value) / (Cells(9, "E").Value)
Cells(16, "E").Value = Application.Sum(Range(Cells(11, "E"), Cells(15, "E")))

ActiveSheet.Range("E1:E24").Columns.AutoFit

End With
End Sub


Sub Button3_Click()
'Inserting a Column at Column B
With ActiveSheet
'Range("E1").EntireColumn.Insert
'Range("E4:E9").Style = "Currency"
'Range("E19:E24").Style = "Currency"
'Range("Q1").EntireColumn.Delete

End With
End Sub
Sub Button4_Click()

With ActiveSheet
'Delete a column at Column B
'Range("E1").EntireColumn.Delete
End With
End Sub
Sub Button1_Click()
'Refresh pivot table data
ActiveWorkbook.RefreshAll
End Sub
Sub Button5_Click()
'get values from CFS AGING SUMMARY


' copy all data from rows E to O to F to P
Worksheets("AGING TRACKING").Range("E3:O24").Copy _
    Destination:=Worksheets("AGING TRACKING").Range("F3:P24")

Worksheets("CFS AGING SUMMARY").Range("J5:J10").Copy Worksheets("AGING TRACKING").Range("E4:E9")
Worksheets("CFS AGING SUMMARY").Range("I5:I10").Copy Worksheets("AGING TRACKING").Range("E19:E24")
With ActiveSheet
' create new percentages
'Cells("E", 11).NumberFormat = "0.00%"

ActiveSheet.Range("E11:E16").NumberFormat = "0.00%"



Cells(11, "E").Value = (Cells(4, "E").Value) / (Cells(9, "E").Value)
Cells(12, "E").Value = (Cells(5, "E").Value) / (Cells(9, "E").Value)
Cells(13, "E").Value = (Cells(6, "E").Value) / (Cells(9, "E").Value)
Cells(14, "E").Value = (Cells(7, "E").Value) / (Cells(9, "E").Value)
Cells(15, "E").Value = (Cells(8, "E").Value) / (Cells(9, "E").Value)
Cells(16, "E").Value = Application.Sum(Range(Cells(11, "E"), Cells(15, "E")))

ActiveSheet.Range("E1:E24").Columns.AutoFit

End With
End Sub

Sub Button3_Click()
'Inserting a Column at Column B
With ActiveSheet
'Range("E1").EntireColumn.Insert
'Range("E4:E9").Style = "Currency"
'Range("E19:E24").Style = "Currency"
'Range("Q1").EntireColumn.Delete

End With
End Sub
Sub Button4_Click()

With ActiveSheet
'Delete a column at Column B
'Range("E1").EntireColumn.Delete
End With
End Sub
Sub Button1_Click()
'Refresh pivot table data
ActiveWorkbook.RefreshAll
End Sub
Sub Button5_Click()
'get values from CFS AGING SUMMARY


' copy all data from rows E to O to F to P
Worksheets("AGING TRACKING").Range("E3:O24").Copy _
    Destination:=Worksheets("AGING TRACKING").Range("F3:P24")

Worksheets("CFS AGING SUMMARY").Range("J5:J10").Copy Worksheets("AGING TRACKING").Range("E4:E9")
Worksheets("CFS AGING SUMMARY").Range("I5:I10").Copy Worksheets("AGING TRACKING").Range("E19:E24")
With ActiveSheet
' create new percentages
'Cells("E", 11).NumberFormat = "0.00%"

ActiveSheet.Range("E11:E16").NumberFormat = "0.00%"



Cells(11, "E").Value = (Cells(4, "E").Value) / (Cells(9, "E").Value)
Cells(12, "E").Value = (Cells(5, "E").Value) / (Cells(9, "E").Value)
Cells(13, "E").Value = (Cells(6, "E").Value) / (Cells(9, "E").Value)
Cells(14, "E").Value = (Cells(7, "E").Value) / (Cells(9, "E").Value)
Cells(15, "E").Value = (Cells(8, "E").Value) / (Cells(9, "E").Value)
Cells(16, "E").Value = Application.Sum(Range(Cells(11, "E"), Cells(15, "E")))

ActiveSheet.Range("E1:E24").Columns.AutoFit

End With
End Sub

Sub Button3_Click()
'Inserting a Column at Column B
With ActiveSheet
'Range("E1").EntireColumn.Insert
'Range("E4:E9").Style = "Currency"
'Range("E19:E24").Style = "Currency"
'Range("Q1").EntireColumn.Delete

End With
End Sub
Sub Button4_Click()

With ActiveSheet
'Delete a column at Column B
'Range("E1").EntireColumn.Delete
End With
End Sub
Sub Button1_Click()
'Refresh pivot table data
ActiveWorkbook.RefreshAll
End Sub
Sub Button5_Click()
'get values from CFS AGING SUMMARY


' copy all data from rows E to O to F to P
Worksheets("AGING TRACKING").Range("E3:O24").Copy _
    Destination:=Worksheets("AGING TRACKING").Range("F3:P24")

Worksheets("CFS AGING SUMMARY").Range("J5:J10").Copy Worksheets("AGING TRACKING").Range("E4:E9")
Worksheets("CFS AGING SUMMARY").Range("I5:I10").Copy Worksheets("AGING TRACKING").Range("E19:E24")
With ActiveSheet
' create new percentages
'Cells("E", 11).NumberFormat = "0.00%"

ActiveSheet.Range("E11:E16").NumberFormat = "0.00%"



Cells(11, "E").Value = (Cells(4, "E").Value) / (Cells(9, "E").Value)
Cells(12, "E").Value = (Cells(5, "E").Value) / (Cells(9, "E").Value)
Cells(13, "E").Value = (Cells(6, "E").Value) / (Cells(9, "E").Value)
Cells(14, "E").Value = (Cells(7, "E").Value) / (Cells(9, "E").Value)
Cells(15, "E").Value = (Cells(8, "E").Value) / (Cells(9, "E").Value)
Cells(16, "E").Value = Application.Sum(Range(Cells(11, "E"), Cells(15, "E")))

ActiveSheet.Range("E1:E24").Columns.AutoFit

End With
End Sub

Sub Button7_Click()


'''MAKE SURE ONLY ONE INSTANCES OF ACCOUNT NUMBER EXISTS
'''MAKE SURE ONLY ONE INSTANCE OF EACH CATEGORY LISTED EXISTS. THIS IS WHAT THE POINT OF THE "FLOYD'S SHEET" TAB IS DOING
'''IT IS SUMMARIZING THE DATA TO MAKE SURE THERE IS NOT TWO VERY SIMILAR ITEMS THAT DO NOT CONSOLIDATE
'''EXAMPLE  "PAYING" VS " PAYING" WHERE THE SECOND ITEM HAS A BEGINNING SPACE



With ActiveSheet
' sets the last row variable to the last row in the first column
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

errorCheck = 0
' loop over all cells '
For i = 2 To (Lastrow)
   
   ' trim spaces on either side of
   Cells(i, "D").Value = Trim(Cells(i, "D").Value)


   ' check row "E" for duplicate account numbers
   currentValue = Cells(i, 5)
   If (currentValue > 1) Then
   
        MsgBox ("duplicate account occured for account number: " & Cells(i, 1) & " in row " & i)
        errorCheck = 1
   
        
    End If
Next i

If (errorCheck = 1) Then
    MsgBox ("Data errors")
    Else: MsgBox ("Data OK")
    End If
End With


End Sub




Sub Button117_Click()


' MAKE SURE NO DATA EXISTS PAST THE LAST ROW OF DATA PASTED
' MAKE SURE FORMULAS WORKING CORRECT

With ActiveSheet
' sets the last row variable to the last row in the first column
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' counts that there is the correct number of rows in the seet
If (Lastrow <> 720) Then
    MsgBox ("The last row is" & Lastrow & "which should be row 720")
    Else: MsgBox ("number of Rows OK")
End If

errorCount = 0

' checks if there are any errors in the sheet
Dim sheetRange As Range
Set sheetRange = Range("A5:DL720")
    
    Dim cell As Range
    
    For Each cell In sheetRange
        errorVal = IsError(cell.Value)
        If (errorVal = True) Then
            errorCount = errorCount + 1
            MsgBox ("Error in cell: " & cell.Address)
         End If

        
        
    Next cell
MsgBox ("Error Count is: " & errorCount)
End With
End Sub

