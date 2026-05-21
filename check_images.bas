Sub CountFigureObjects()

    Dim inlineCount As Long
    Dim shapeCount As Long
    Dim totalCount As Long

    ' Картинки внутри текста
    inlineCount = ActiveDocument.InlineShapes.Count

    ' Плавающие объекты
    shapeCount = ActiveDocument.Shapes.Count

    totalCount = inlineCount + shapeCount

    MsgBox "InlineShapes: " & inlineCount & vbCrLf & _
           "Shapes: " & shapeCount & vbCrLf & _
           "Всего рисунков: " & totalCount, _
           vbInformation

End Sub
