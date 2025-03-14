Sub UpdateDocumentProperties()
    Application.ScreenUpdating = False
    Application.Options.UpdateFieldsAtPrint = False
    'On Error GoTo ErrorHandler

    With ActiveDocument
        ' Update each custom document property
        .CustomDocumentProperties("BOK ID").Value = "New BOK ID"
        .CustomDocumentProperties("Document Name").Value = "New Document Name"
        .CustomDocumentProperties("Company Name").Value = "New Company Name"
        .CustomDocumentProperties("Division").Value = "New Division"
        .CustomDocumentProperties("Author").Value = "New Author"
        .CustomDocumentProperties("Company Address").Value = "New Company Address"
        .CustomDocumentProperties("Project Name").Value = "New Project Name"
        .CustomDocumentProperties("Project Number").Value = "New Project Number"
        .CustomDocumentProperties("End Customer").Value = "New End Customer"
        .CustomDocumentProperties("Site Name").Value = "New Site Name"
        .CustomDocumentProperties("File Name").Value = "New File Name"

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
    
    UpdateTitlePageFields
    
    ActiveDocument.Repaginate

    Dim TOC As TableOfContents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Update
    Next

    ' Inform the user
    MsgBox "Properties updated successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.Options.UpdateFieldsAtPrint = True
    Exit Sub

'ErrorHandler:
    'MsgBox "Error updating properties: " & Err.Description, vbCritical
End Sub

Sub UpdateTitlePageFields()
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