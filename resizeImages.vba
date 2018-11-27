Sub resizePics()
'
' resizePics Macro
' 
' This resizes all photos taller than desiredSize to a a height of 
' desiredSize while maintaining the original aspect ratio
'
    Dim s As Single
    Dim i As Long
    Dim newHeight As Integer
    Dim newWidth As Integer
    Dim desiredSize As Integer
    
    desiredSize = 8
    
    With ActiveDocument
        For i = 1 To .InlineShapes.Count
            With .InlineShapes(i)
                s = Application.CentimetersToPoints(desiredSize) / .Height

                If s < 1 Then
                    newHeight = .Height * s
                    newWidth = .Width * s
                    .Height = newHeight
                    .Width = newWidth
                    Debug.Print ("Modifying image " & i)
                End If
            End With
        Next i
    End With

End Sub