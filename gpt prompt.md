
# Edit or rewrite this code.

The following is a vba macro for a word document that contains the following custom properties:
"BOK ID"
"Document Name"
"Company Name"
"Division"
"Author"
"Company Address"
"Project Name"
"Project Number"
"End Customer"
"Site Name"
"File Name"

all valus are text values.

The following code is a part of a VBA form.
This is the code of the submittal button once the form is filled out.

```vb
Private Sub UpdatePropertiesButton_Click()
    Application.ScreenUpdating = False
    Application.Options.UpdateFieldsAtPrint = False
    'On Error GoTo ErrorHandler

    With ActiveDocument
        ' Update each custom document property
        .CustomDocumentProperties("BOK ID").Value = txtBOKID.Value
        .CustomDocumentProperties("Document Name").Value = txtDocumentName.Value
        .CustomDocumentProperties("Company Name").Value = txtCompanyName.Value
        .CustomDocumentProperties("Division").Value = txtDivision.Value
        .CustomDocumentProperties("Author").Value = txtAuthor.Value
        .CustomDocumentProperties("Company Address").Value = txtCompanyAddress.Value
        .CustomDocumentProperties("Project Name").Value = txtProjectName.Value
        .CustomDocumentProperties("Project Number").Value = txtProjectNumber.Value
        .CustomDocumentProperties("End Customer").Value = txtEndCustomer.Value
        .CustomDocumentProperties("Site Name").Value = txtSiteName.Value
        .CustomDocumentProperties("File Name").Value = txtFileName.Value

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

    ' Close the form if needed
    Unload Me

    Application.ScreenUpdating = True
    Application.Options.UpdateFieldsAtPrint = True
    Exit Sub

'ErrorHandler:
    'MsgBox "Error updating properties: " & Err.Description, vbCritical
End Sub

Private Sub UpdateTitlePageFields()
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



Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Caption = "Update Document Properties"
    txtBOKID.Value = ActiveDocument.CustomDocumentProperties("BOK ID").Value
    txtDocumentName.Value = ActiveDocument.CustomDocumentProperties("Document Name").Value
    txtCompanyName.Value = ActiveDocument.CustomDocumentProperties("Company Name").Value
    txtDivision.Value = ActiveDocument.CustomDocumentProperties("Division").Value
    txtAuthor.Value = ActiveDocument.CustomDocumentProperties("Author").Value
    txtCompanyAddress = ActiveDocument.CustomDocumentProperties("Company Address").Value
    txtProjectName = ActiveDocument.CustomDocumentProperties("Project Name").Value
    txtProjectNumber = ActiveDocument.CustomDocumentProperties("Project Number").Value
    txtEndCustomer = ActiveDocument.CustomDocumentProperties("End Customer").Value
    txtSiteName = ActiveDocument.CustomDocumentProperties("Site Name").Value
    txtFileName = ActiveDocument.CustomDocumentProperties("File Name").Value
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

The new code must now run as a macro outside of the form so an outside program can target and exicute the macro using a cscript.

The outside program is expected to edit the word document custom property values
"BOK ID"
"Document Name"
"Company Name"
"Division"
"Author"
"Company Address"
"Project Name"
"Project Number"
"End Customer"
"Site Name"
"File Name"

externally, then launch the word document and exicute the macro so that the new property values get filled into document.

Edit the code so that it can be called externally by cscript.
