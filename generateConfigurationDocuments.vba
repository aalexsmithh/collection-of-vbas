Sub generateConfigurationDocuments()
'
' generateConfigurationDocuments Macro
' 
' This generates a configuration document from each line of the configuration tracker
' Screenshots are pulled from the corresponding image folder in the root of the configuration trakcer excel file
'

	Dim sh As Worksheet
	Dim rw As Range
	Dim rowCount As Integer

	Set sh = ActiveSheet
	For Each rw in sh.Rows
		' Check if we are at the end of the document
		If sh.Cells(rw.Row, 1) = "" Then
			Exit For
		End If

		

	End For