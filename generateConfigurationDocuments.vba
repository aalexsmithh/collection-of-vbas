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
	Dim s As Single
	Dim newHeight As Integer
	Dim newWidth As Integer
	Dim desiredSize As Integer
	Dim useableDocWidth As Integer

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
				useableDocWidth = .PageSetup.PageWidth - .PageSetup.LeftMargin - .PageSetup.RightMargin
				.Windows(1).Visible = True

				' Create table, add borders, and set width of columns
				.Tables.Add Range:=.Range(0,0), NumRows:=8, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior 
				currTable = .Tables(1)
				currTable.Borders.Enable = True
				currTable.Columns(1).Borders.Enable = True
				currTable.Columns(1).SetWidth ColumnWidth:=108, RulerStyle:=wdAdjustProportional
				
				' Populate the table with information from the excel sheet
				rowCount = 1
				For each wordTableRow in currTable.rows
					colCount = 1
					For each wordTableCell in wordTableRow.cells
						If colCount = 1 Then
							wordTableCell.Range.InsertAfter sheet.Cells(1,rowCount)
						Else
							wordTableCell.Range.InsertAfter sheet.Cells(row.Row, rowCount)
						End If
						colCount = colCount + 1
					Next
					rowCount = rowCount + 1
				Next

				' Pull the screenshots from each folder and add them to the document
				Set fileSearcher = CreateObject("Scripting.FileSystemObject")
				For Each file In fileSearcher.GetFolder(Application.ActiveWorkbook.Path & "/" & sheet.Cells(row.Row, 1)).Files
					On Error Goto 0
					Dim img As Object
					' .Content.InsertAfter
					' .Characters.Last.Select
					' Selection.Collapse

					Set img = .Characters.Last.InlineShapes.AddPicture(file.Path)
					resizeImage img:=img, docWidth:=useableDocWidth

				Next file
				' .Content.InsertAfter sheet.Cells(1, 1)

				.SaveAs(Application.ActiveWorkbook.Path & "/" & sheet.Cells(row.Row, 1) & ".docx")
				.Close
				Exit For
			End With
		End If
		
		'Exit For
	Next row


End Sub

Sub resizeImage(img, docWidth)
'
' requires the image to resize as well as the document width (in points)
' will resize the image if it is wider than the document, 
' otherwise it will remain the same
'

	Dim newHeight As Integer
	Dim newWidth As Integer
	Dim s As Single

	' Debug.Print ("img width " & img.Width)
	If img.Width > docWidth Then
		s = Application.CentimetersToPoints(docWidth) / img.Width
		newHeight = img.Height * s
		newWidth = img.Width * s
		img.Height = newHeight
		img.Width = newWidth
		Debug.Print ("Modifying image " & i)
	End If

End Sub
