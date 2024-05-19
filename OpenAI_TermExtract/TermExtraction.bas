Attribute VB_Name = "TermExtraction"
Sub ProcessTermsAndCopySplitTerms()
    Dim wsTerms As Worksheet
    Dim wsSource As Worksheet
    Dim wsDefinitions As Worksheet
    Dim srcData As String
    Dim lines As Variant
    Dim line As Variant
    Dim rowIndex As Integer
    Dim dataRow As Integer
    Dim lastRow As Integer

    ' Check if "Terms" worksheet exists
    Set wsTerms = ThisWorkbook.Sheets("Terms")
    If wsTerms Is Nothing Then
        MsgBox "The 'Terms' worksheet does not exist. Please check and try again."
        Exit Sub
    End If
    
    ' Apply formula in Terms sheet and replace with values
    lastRow = wsTerms.Cells(wsTerms.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        wsTerms.Cells(i, 3).Formula = "=CallOpenAI_prompt(A" & i & " & ""~"" & B" & i & ", Prompts!$A$3)"
    Next i
    wsTerms.Calculate
    For i = 1 To lastRow
        wsTerms.Cells(i, 3).Value = wsTerms.Cells(i, 3).Value
    Next i

    ' Set the source worksheet
    Set wsSource = ThisWorkbook.ActiveSheet ' Use the active sheet or specify if needed

    ' Check if "Definitions" worksheet exists
    On Error Resume Next
    Set wsDefinitions = ThisWorkbook.Sheets("Definitions")
    On Error GoTo 0

    ' Create "Definitions" sheet if it doesn't exist
    If wsDefinitions Is Nothing Then
        Set wsDefinitions = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDefinitions.Name = "Definitions"
    Else
        wsDefinitions.Cells.Clear ' Clear existing data if sheet exists
    End If

    ' Define headers
    wsDefinitions.Cells(1, 1).Value = "Source Term"
    wsDefinitions.Cells(1, 2).Value = "Target Term"
    wsDefinitions.Cells(1, 3).Value = "Output"  ' Additional header for the formula results

    ' Initialize row index for Definitions sheet
    rowIndex = 2

    ' Process each row in the source sheet that contains data
    For dataRow = 1 To wsSource.Cells(wsSource.Rows.Count, 3).End(xlUp).Row
        srcData = wsSource.Cells(dataRow, 3).Value
        lines = Split(srcData, Chr(10)) ' Split data into lines

        For Each line In lines
            line = Split(line, "|") ' Split each line at the pipe symbol
            If UBound(line) >= 1 Then ' Ensure there are two parts
                wsDefinitions.Cells(rowIndex, 1).Value = Trim(line(0)) ' Copy first part
                wsDefinitions.Cells(rowIndex, 2).Value = Trim(line(1)) ' Copy second part
                ' Insert the formula into column C, adjusting for the current row
                wsDefinitions.Cells(rowIndex, 3).Formula = "=CallOpenAI_prompt(A" & rowIndex & ", Prompts!$B$3)"
                rowIndex = rowIndex + 1
            End If
        Next line
    Next dataRow

    ' Evaluate formulas and replace them with their results in Definitions
    wsDefinitions.Calculate
    For i = 2 To rowIndex - 1
        wsDefinitions.Cells(i, 3).Value = wsDefinitions.Cells(i, 3).Value
    Next i

    ' Autofit columns for better visibility in Definitions
    wsDefinitions.Columns("A:C").AutoFit

    MsgBox "Processing complete for both Terms and Definitions sheets."
End Sub

