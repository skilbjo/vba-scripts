

Sub CompilationDF()
    Dim Worksheets As Integer, i, endRow, rowCount
    Dim PacketWorkbook As Workbook
    Dim resultsWorkbook As Workbook
    Dim file As String
    Dim sourceRange As Range, destRange
    
    file = ActiveWorkbook.Worksheets("VBA").Range("B3").Value
    endRow = 1
    rowCount = 37 + 1
    Set resultsWorkbook = ActiveWorkbook
    Set PacketWorkbook = Workbooks.Open(file)
    Worksheets = PacketWorkbook.Worksheets.Count
    
    
    For i = 1 To Worksheets
        ' Set sourceRange
        If PacketWorkbook.Worksheets(i).Name Like "DF*" Then
            Set sourceRange = PacketWorkbook.Worksheets(i).Range("A9:E51")
            Set destRange = resultsWorkbook.Worksheets("ResultsDF").Range("a" & endRow & ":e" & endRow + rowCount)
            Set destRange = destRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
            destRange.Value = sourceRange.Value
            ' Increment
            endRow = endRow + rowCount
        End If
    ' Next tab
    Next i
    
    MsgBox "We're done here"

End Sub


Sub CompilationB()
    Dim Worksheets As Integer, i, endRow, rowCount
    Dim PacketWorkbook As Workbook
    Dim resultsWorkbook As Workbook
    Dim file As String
    Dim sourceRange As Range, destRange
    
    file = ActiveWorkbook.Worksheets("VBA").Range("B3").Value
    endRow = 1
    rowCount = 116 + 1
    Set resultsWorkbook = ActiveWorkbook
    Set PacketWorkbook = Workbooks.Open(file)
    Worksheets = PacketWorkbook.Worksheets.Count
    
    
    For i = 1 To Worksheets
        ' Set sourceRange
        If PacketWorkbook.Worksheets(i).Name Like "B*" Then
            Set sourceRange = PacketWorkbook.Worksheets(i).Range("A8:E136")
            Set destRange = resultsWorkbook.Worksheets("ResultsB").Range("a" & endRow & ":e" & endRow + rowCount)
            Set destRange = destRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
            destRange.Value = sourceRange.Value
            ' Increment
            endRow = endRow + rowCount
        End If
    ' Next tab
    Next i
    
    MsgBox "We're done here"

End Sub

