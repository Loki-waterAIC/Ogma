
Sub ogmaMacro()
    Dim doc As Document
    Set doc = ActiveDocument

    Application.ScreenUpdating = False
    Application.Options.UpdateFieldsAtPrint = False

    With ActiveDocument

        ' Update fields in the main document body
        .Fields.Update

        ' Update fields in headers and footers for each section
        Dim oSection As Section
        Dim oHeader As HeaderFooter
        Dim oFooter As HeaderFooter
        For Each oSection In .Sections
            oSection.Range.Fields.Update
            For Each oHeader In oSection.Headers
                oHeader.Range.Fields.Update
            Next oHeader
            For Each oFooter In oSection.Footers
                oFooter.Range.Fields.Update
            Next oFooter
        Next oSection
    End With

    ogmaUpdateTitlePageFields
    
    ActiveDocument.Repaginate

    Dim TOC As TableOfContents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Update
    Next


    Application.ScreenUpdating = True
    Application.Options.UpdateFieldsAtPrint = True
End Sub

Private Sub ogmaUpdateTitlePageFields()
    Dim oShape As Shape
    Dim iPageNumber As Integer
    Dim oRange As Range

    iPageNumber = 1 ' title page

    For Each oShape In ActiveDocument.Shapes
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
End Sub
