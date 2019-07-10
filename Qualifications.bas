Private Sub processFiles(folder)

    Dim file As String
    Dim wb As Workbook
    
    If Right(folder, 1) <> "\" Then
        folder = folder & "\"
    End If
    
    file = Dir(folder & "DOFORMS*")

    Do While Len(file) > 0
        Set wb = Workbooks.Open(folder & file)
        doWork wb
        'wb.Close SaveChanges:=True
        file = Dir()
    Loop
    
End Sub
Private Sub doWork(wb)

    Dim i As Integer
    Dim a As Integer
    Dim codes(1 To 100) As String

    'insert headers
    Range("A3").Value = "LABORCODE"
    Range("B3").Value = "ORGID"
    Range("C3").Value = "WORKSITE"
    Range("D3").Value = "LABORQUAL.QUALIFICATIONID"
    Range("E3").Value = "LABORQUAL.CERTIFICATENUM"
    Range("F3").Value = "LABORQUAL.EFFDATE"
    Range("G3").Value = "LABORQUAL.VALIDATIONDATE"
    Range("H3").Value = "LABORQUAL.STATUS"
    
    'find labor codes
    i = 1
    a = 1
    Do While Cells(1, i).Value <> ""
        If Left(Cells(1, i), 4) = "Code" And Cells(2, i).Value <> "" Then
            codes(a) = Cells(2, i).Value
            a = a + 1
        End If
        i = i + 1
    Loop
    
    'add labor codes
    i = 4
    Do While i < a + 5
        Cells(i, 1).Value = codes(i - 3)
        i = i + 1
    Loop
    i = i - 3
    Range("A4:A" & i).NumberFormat = "0000"
    
    'add other values
    Range("B4:B" & i).Value = "31"
    Range("C4:C" & i).Value = "1000"
    Range("D4:D" & i).Value = Range("F2").Value
    Range("E4:E" & i).Value = "=CONCATENATE(RC[-1],""."",TEXT(RC[-4],""0000""))"
    Range("F4:G" & i).Value = DateAdd("d", 1, Range("H2").Value)
    Range("F4:G" & i).NumberFormat = "yyyy-mm-dd"
    Range("H4:H" & i).Value = "ACTIVE"
    
    'delete top rows
    Rows(1).Delete
    Rows(1).Delete
    
End Sub
Sub DoForms()

    Dim folder As String
    
    folder = InputBox("Enter folder path (Ex. 'H:\DoForms'): ")
    
    processFiles folder
    
End Sub
