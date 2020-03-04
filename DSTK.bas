Attribute VB_Name = "Petes_Data_Science_Toolkit"
Option Explicit

Sub export_in_json_format()

    Dim fs As Object
    Dim jsonfile
    Dim rangetoexport As Range
    Dim rowcounter As Long
    Dim columncounter As Long
    Dim linedata As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Set jsonfile = fs.CreateTextFile(ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".js", True)
    
    linedata = "var " & ActiveSheet.Name & " = ["
    jsonfile.WriteLine linedata
    For rowcounter = 2 To ActiveSheet.UsedRange.Rows.Count
        linedata = ""
        For columncounter = 1 To ActiveSheet.UsedRange.Columns.Count
            linedata = linedata & """" & Replace(Cells(1, columncounter).Value, """", "'") & """" & ":" & """" & Replace(Cells(rowcounter, columncounter).Value, """", "'") & """" & ","
        Next
        linedata = Left(linedata, Len(linedata) - 1)
        If rowcounter = ActiveSheet.UsedRange.Rows.Count Then
            linedata = "{" & linedata & "}"
        Else
            linedata = "{" & linedata & "},"
        End If
        
        jsonfile.WriteLine linedata
    Next
    linedata = "];"
    jsonfile.WriteLine linedata
    jsonfile.Close
    
    Set fs = Nothing
    
    
End Sub


Sub SelectionToCSV()
    
    Dim myFile As String, rng As Range, cellValue As Variant, i As Integer, j As Integer
    
    myFile = Application.DefaultFilePath & "\sales.csv"
    
    Set rng = Selection
    
    Open myFile For Output As #1
    
    For i = 1 To rng.Rows.Count
         For j = 1 To rng.Columns.Count
    
            cellValue = rng.Cells(i, j).Value
    
            cellValue = rng.Cells(i, j).Value
    
        Next j
    Next i
    
    Close #1

End Sub


Function FileToString(filelocation As String)

 ' file to string

    Dim strFilename As String: strFilename = filelocation ' something like "C:\temp\yourfile.txt"
    Dim strFileContent As String
    Dim iFile As Integer: iFile = FreeFile
    Open strFilename For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    FileToString = strFileContent

End Function


Function RowsOnSheet(SheetName As String)
    Dim i As Integer
    i = Cells(Sheets(SheetName).Rows.Count, 1).End(xlUp).Row
    RowsOnSheet = i
End Function

Function DTVLookup(TheValue As Variant, TheRange As Range, TheColumn As Long, Optional PercentageMatch As Double = 100) As Variant 'uses levenshtein3 function to perform a fuzzy Vlookup If TheColumn < 1 Then
    DTVLookup = CVErr(xlErrValue)
    Exit Function
End If
If TheColumn > TheRange.Columns.Count Then
    DTVLookup = CVErr(xlErrRef)
    Exit Function
End If
Dim c As Range
For Each c In TheRange.Columns(1).Cells
    If UCase(TheValue) = UCase(c) Then
        DTVLookup = c.Offset(0, TheColumn - 1)
        Exit Function
    ElseIf PercentageMatch <> 100 Then
        If Levenshtein3(UCase(TheValue), UCase(c)) >= PercentageMatch Then
            DTVLookup = c.Offset(0, TheColumn - 1)
            Exit Function
        End If
    End If
Next c
DTVLookup = CVErr(xlErrNA)
End Function

Function Levenshtein3(ByVal string1 As String, ByVal string2 As String) As Long 'returns a percent match between strings

Dim i As Long, j As Long, string1_length As Long, string2_length As Long
Dim distance(0 To 60, 0 To 50) As Long, smStr1(1 To 60) As Long, smStr2(1 To 50) As Long
Dim min1 As Long, min2 As Long, min3 As Long, minmin As Long, MaxL As Long

string1_length = Len(string1):  string2_length = Len(string2)

distance(0, 0) = 0
For i = 1 To string1_length:    distance(i, 0) = i: smStr1(i) = Asc(LCase(Mid$(string1, i, 1))): Next
For j = 1 To string2_length:    distance(0, j) = j: smStr2(j) = Asc(LCase(Mid$(string2, j, 1))): Next
For i = 1 To string1_length
    For j = 1 To string2_length
        If smStr1(i) = smStr2(j) Then
            distance(i, j) = distance(i - 1, j - 1)
        Else
            min1 = distance(i - 1, j) + 1
            min2 = distance(i, j - 1) + 1
            min3 = distance(i - 1, j - 1) + 1
            If min2 < min1 Then
                If min2 < min3 Then minmin = min2 Else minmin = min3
            Else
                If min1 < min3 Then minmin = min1 Else minmin = min3
            End If
            distance(i, j) = minmin
        End If
    Next
Next

' Levenshtein3 will properly return a percent match (100%=exact) based on similarities and Lengths etc...
MaxL = string1_length: If string2_length > MaxL Then MaxL = string2_length
Levenshtein3 = 100 - CLng((distance(string1_length, string2_length) * 100) / MaxL)

End Function

Sub export_rows_to_docs()

    Dim fs As Object
    Dim docfile
    Dim rangetoexport As Range
    Dim rowcounter As Long
    Dim columncounter As Long
    Dim linedata As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Set docfile = fs.CreateTextFile(ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".html", True)
    
    linedata = ""
    docfile.WriteLine linedata
    For rowcounter = 2 To ActiveSheet.UsedRange.Rows.Count
        linedata = "###<h3>" & Cells(rowcounter, 2).Value & ", " & Cells(rowcounter, 3).Value & " (" & Cells(rowcounter, 1).Value & ")</h3>"
        For columncounter = 1 To ActiveSheet.UsedRange.Columns.Count
            If Len(Cells(rowcounter, columncounter).Value) > 0 Then
                linedata = linedata & "<p><b>" & Cells(1, columncounter).Value & "</b> : " & Cells(rowcounter, columncounter).Value & "<br><br>"
            End If
        Next
        docfile.WriteLine linedata
    Next
    docfile.WriteLine linedata
    docfile.Close
    
    Set fs = Nothing
    
End Sub

Sub CreateSheet(SheetName As String)
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = SheetName
End Sub

Sub Columnify_Table()
    Application.ScreenUpdating = False

    Dim oldsheet As Variant
    oldsheet = ActiveWorkbook.ActiveSheet.Name

    CreateSheet ("temp")

    Sheets(oldsheet).Activate
    
    Dim R, i, j As Double
    
    Dim X, Y, Z As Variant
    
    R = 2

    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        For j = 1 To ActiveSheet.UsedRange.Columns.Count
            X = Cells(i, 1).Value
            Y = Cells(1, j).Value
            Z = Cells(i, j).Value
            Sheets("temp").Activate
            Cells(R, 1).Value = X
            Cells(R, 2).Value = Y
            Cells(R, 3).Value = Z
            R = R + 1
            Sheets(oldsheet).Activate
        Next j
    Next i

    Application.ScreenUpdating = True

End Sub
