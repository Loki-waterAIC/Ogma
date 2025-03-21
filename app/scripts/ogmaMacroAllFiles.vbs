Sub ogmaMacroAllFiles()
    Dim doc As Document
    Dim TOC As TableOfContents
    Dim oSection As Section
    Dim oHeader As HeaderFooter
    Dim oFooter As HeaderFooter
    Dim oShape As Shape
    Dim iPageNumber As Integer
    Dim oRange As Range

    Application.ScreenUpdating = False
    Application.Options.UpdateFieldsAtPrint = False

    ' Loop through all open documents
    For Each doc In Application.Documents
        With doc
            ' Update fields in the main document body
            .Fields.Update

            ' Update fields in headers and footers for each section
            For Each oSection In .Sections
                oSection.Range.Fields.Update
                For Each oHeader In oSection.Headers
                    oHeader.Range.Fields.Update
                Next oHeader
                For Each oFooter In oSection.Footers
                    oFooter.Range.Fields.Update
                Next oFooter
            Next oSection

            ' Update title page fields
            iPageNumber = 1 ' title page
            For Each oShape In .Shapes
                ' Check if the shape is on the first page
                If oShape.Anchor.Information(wdActiveEndAdjustedPageNumber) = iPageNumber Then
                    ' Check if the shape has a text frame
                    If Not oShape.TextFrame Is Nothing Then
                        ' Additional check: Make sure the text frame has text
                        If oShape.TextFrame.HasText Then
                            Set oRange = oShape.TextFrame.TextRange
                            oRange.Fields.Update
                        End If
                    End If
                End If
            Next oShape

            ' Repaginate the document
            .Repaginate

            ' Update table of contents
            For Each TOC In .TablesOfContents
                TOC.Update
            Next TOC
        End With
    Next doc

    Application.ScreenUpdating = True
    Application.Options.UpdateFieldsAtPrint = True
End Sub