Sub legal_activity_check()


'''MAKE SURE ONLY ONE INSTANCES OF ACCOUNT NUMBER EXISTS	
'''MAKE SURE ONLY ONE INSTANCE OF EACH CATEGORY LISTED EXISTS. THIS IS WHAT THE POINT OF THE "FLOYD'S SHEET" TAB IS DOING	
'''IT IS SUMMARIZING THE DATA TO MAKE SURE THERE IS NOT TWO VERY SIMILAR ITEMS THAT DO NOT CONSOLIDATE
'''EXAMPLE  "PAYING" VS " PAYING" WHERE THE SECOND ITEM HAS A BEGINNING SPACE



With ActiveSheet
' sets the last row variable to the last row in the first column
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' loop over all cells '
For i = 2 To (lastRow)
   
   ' trim spaces on either side of
   Cells(i, "D").Value = Trim(Cells(i, "D").Value)


   ' check row "E" for duplicate account numbers
   currentValue = Cells(i, 5)
   If (currentValue > 1) Then
   
        MsgBox ("duplicate account occured for account number: " & Cells(i, 1) & " in row " & i)
        
    End If
Next i

End With
End Sub

