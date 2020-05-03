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
If (lastRow <> 720) Then
    MsgBox ("The last row is" & lastRow & "which should be row 720")
    Else: MsgBox ("number of Rows OK")
End If

errorCount = 0

Dim sheetRange As Range
Set sheetRange = Range("A5:DL720")
    
    Dim cell As Range
    
    For Each cell In sheetRange
        errorVal = IsError(cell.Value)
        If (errorVal = True) Then
            errorCount = errorCount + 1
         End If

        
        
    Next cell
MsgBox ("Error Count is: " & errorCount)
End With
End Sub
