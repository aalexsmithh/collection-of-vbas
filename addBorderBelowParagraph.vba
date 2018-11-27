Sub addBorderBelowParagraph()
'
' addBorderBelowParagraph Macro
' finds all paragraphs of style styleToActOn and applies a border
' under those paragraphs
'
	Dim styleToActOn As String

	styleToActOn = "Heading 2"

	For Each para In ActiveDocument.Paragraphs
		' Debug.Print (para.Style)
		If para.Style = styleToActOn Then
			With para.Borders(wdBorderBottom)
				.LineStyle = wdLineStyleSingle
			End With
		End If
			
	Next para

End Sub