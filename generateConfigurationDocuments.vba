Sub generateConfigurationDocuments()
'
' generateConfigurationDocuments Macro
' 
' This generates a configuration document from each line of the configuration tracker
' Screenshots are pulled from the corresponding image folder in the root of the configuration trakcer excel file
'

	Dim sheet As Worksheet
	Dim row As Range
	Dim rowCount As Integer
	Dim colCount As Integer

	Set sheet = ActiveSheet
	For Each row in sheet.Rows
		' Check if we are at the end of the document
		If sheet.Cells(row.Row, 1) = "" Then
			Exit For
		End If

		' Excel Table fields
		' [ID,Name,Transport Request,Request Type,Customizing Task,Path,Cross-client?,Transaction]

		If row.Row > 1 Then

			With CreateObject("Word.Document")
				.Windows(1).Visible = True

				' Create table and add borders
				currTable = .Tables.Add Range:=.Range(0,0), NumRows:=8, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior 
				' .Tables(1)
				currTable.Borders.Enable = True
				currTable.Columns(1).Width = 10

				rowCount = 1
				For each wordTableRow in currTable.rows
					colCount = 1
					For each wordTableCell in wordTableRow.cells
						If colCount = 1 Then
							wordTableCell.Range.InsertAfter sheet.Cells(1,rowCount)
							' wordTableCell.Borders(wdBorderRight).Enable = True
						Else
							wordTableCell.Range.InsertAfter sheet.Cells(row.Row, rowCount)
						End If
						colCount = colCount + 1
					Next
					rowCount = rowCount + 1
				Next

				.Content.InsertAfter sheet.Cells(1, 1)

				.SaveAs(Application.ActiveWorkbook.Path & "/" & sheet.Cells(row.Row, 1) & ".docx")
				.Close
				Exit For
			End With
		End If
		
		'Exit For
	Next row


	

End Sub