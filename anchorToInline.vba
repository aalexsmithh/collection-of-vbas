Sub anchorToInline()
'
' anchorToInline Macro
' Resets all images to be inline instead of being anchored
' with text wrapping.
'

    Dim i As Integer, oShp As Shape

    For i = ActiveDocument.Shapes.Count To 1 Step -1
        ' Select the document
        With ActiveDocument.Shapes(i)
            ' Cut and paste the image as inline
            .Select
            Selection.Cut
            Selection.PasteSpecial dataType:=wdPasteBitmap, Placement:=wdInLine
        End With
    Next i

End Sub