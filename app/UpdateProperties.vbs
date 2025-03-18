' Make sure this VBA script is inside the normal.dot/.dotm file.
' You can add via in Micorosoft Word > Developer tab (needs to be enabled in the ribbon to see) >
'   Visual Basic > Normal (Right Click) (located on the top left of the left hand side navigation menu) > 
'   Insert > Module.
' Paste in the script. To name the scipt, go into the module properties and and change the file name.
' Macros are called by their functions, not their file names.

Sub UpdateProperties()
    Application.ScreenUpdating = False
    Application.Options.UpdateFieldsAtPrint = False
    'On Error GoTo ErrorHandler

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
