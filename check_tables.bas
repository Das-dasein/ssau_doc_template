Sub CountTablesByCaptions_Strict()

    Dim p As Paragraph
    Dim txt As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Dim totalTables As Long
    Dim part As String

    totalTables = 0

    Set re = CreateObject("VBScript.RegExp")

    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "Таблица\s+[0-9]+(\s*[-–]\s*[0-9]+)?"
    End With

    For Each p In ActiveDocument.Paragraphs

        txt = p.Range.Text

        If re.Test(txt) Then

            Set matches = re.Execute(txt)

            For Each m In matches
                totalTables = totalTables + 1
            Next m

        End If

    Next p

    MsgBox "Всего таблиц (по подписям): " & totalTables, vbInformation

End Sub
