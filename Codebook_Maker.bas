Attribute VB_Name = "Module1"
Sub BuildCodebook()

    cols = ActiveSheet.UsedRange.Columns.Count
    rws = ActiveSheet.UsedRange.Rows.Count

    

    tempCodebook = "<html><head><style>td{border-left:1px solid black;border-top:1px solid black;}Table{border-right:1px solid black;border-bottom:1px solid black;}</style></head><body><h3>" + ActiveSheet.Name + "</h3><p>Record Type is: Undefined</p><p> Number of Records: " + Str(rws - 1) + "</p><table>"

    tempSet = ""
    
    
    
    For i = 1 To cols
        tempSet = "<br><br>"
        tempQuestion = Cells(1, i).Value
        tempNumber = i
        tempCardinality = 0
        For j = 2 To rws
            
            
            If IsDate(Cells(j, i).Value) And Len(Cells(j, i).Value) <> 0 Then
                 tempSet = "Is Date"
                Exit For
            ElseIf IsNumeric(Cells(j, i).Value) And Len(Cells(j, i).Value) <> 0 Then
                tempSet = "Is Numeric"
                Exit For
            ElseIf tempCardinality > 30 Then
                tempSet = "Open Ended Text (>30 values)"
                Exit For
            Else
                If Len(Cells(j, i).Value) <> 0 And InStr(tempSet, "<br>" + Cells(j, i).Value) = 0 Then
                    tempSet = tempSet + "<br>" + Cells(j, i).Value
                    tempCardinality = tempCardinality + 1
                End If
            End If
        Next j
        tempCodebook = tempCodebook + "<tr>" + "<td>" + Str(tempNumber) + "</td><td>" + tempQuestion + "</td><td>" + tempSet + "</td><td>Description</td></tr>" + Chr(10)
    Next i
        
    tempCodebook = tempCodebook + "</table></body></html>"
    
    FilePath = ActiveWorkbook.Path & "\codebook" + Format(Now, "yyyymmddHHmmss") + ".html"
    Open FilePath For Output As #2
    Write #2, tempCodebook
    Close #2
    

End Sub
