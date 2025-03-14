Sub UpdateAllFieldsAndCustomProperties()
    Dim doc As Document
    Set doc = ActiveDocument
    
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
        
        ' Update custom properties
        UpdateCustomProperties doc
    End With
End Sub

Sub UpdateCustomProperties(doc As Document)
    Dim propNames As Variant
    Dim propValues As Variant
    Dim i As Integer
    
    ' Define the custom properties and their values
    propNames = Array("BOK ID", "Document Name", "Company Name", "Division", "Author", "Company Address", "Project Name", "Project Number", "End Customer", "Site Name", "File Name")
    propValues = Array("12345", "Sample Document", "My Company", "My Division", "John Doe", "1234 Elm St", "Project X", "98765", "Customer Y", "Site Z", "Document.docx")
    
    ' Update each custom property
    For i = LBound(propNames) To UBound(propNames)
        doc.CustomDocumentProperties(propNames(i)).Value = propValues(i)
    Next i
End Sub