Sub CheckReferenceOrder()

    Dim rng As Range
    Dim searchRange As Range

    Dim refText As String
    Dim parts() As String
    Dim part As Variant

    Dim prevNum As Long
    Dim num As Long
    Dim startNum As Long
    Dim endNum As Long
    Dim i As Long

    Dim errors As String
    Dim appendixPos As Long

    Dim totalRefs As Long

    prevNum = 0
    errors = ""
    totalRefs = 0

    Set searchRange = ActiveDocument.Range

    appendixPos = InStr(1, UCase(searchRange.Text), "ПРИЛОЖЕНИЕ")

    If appendixPos > 0 Then
        Set searchRange = ActiveDocument.Range(0, appendixPos - 1)
    End If

    Set rng = searchRange.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "\[[0-9,\- ]@\]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute

        refText = rng.Text

        ' убрать скобки
        refText = Replace(refText, "[", "")
        refText = Replace(refText, "]", "")

        parts = Split(refText, ",")

        For Each part In parts

            part = Trim(part)

            ' диапазон 1-5
            If InStr(part, "-") > 0 Then

                startNum = CLng(Split(part, "-")(0))
                endNum = CLng(Split(part, "-")(1))

                For i = startNum To endNum

                    totalRefs = totalRefs + 1

                    If i < prevNum Then
                        errors = errors & _
                            "Ошибка: " & i & _
                            " после " & prevNum & vbCrLf
                    End If

                    prevNum = i

                Next i

            Else

                num = CLng(part)

                totalRefs = totalRefs + 1

                If num < prevNum Then
                    errors = errors & _
                        "Ошибка: " & num & _
                        " после " & prevNum & vbCrLf
                End If

                prevNum = num

            End If

        Next part

        rng.Collapse wdCollapseEnd

    Loop

    If errors = "" Then
        MsgBox "Порядок ссылок корректный." & vbCrLf & _
               "Всего ссылок: " & totalRefs, vbInformation
    Else
        MsgBox "Найдены ошибки:" & vbCrLf & vbCrLf & _
               errors & vbCrLf & _
               "Всего ссылок: " & totalRefs, vbExclamation
    End If

End Sub
