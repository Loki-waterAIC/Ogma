' single file
Sub AutoOpen()
    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        fld.Update
    Next
End Sub

' all opened files
Sub allOpenedFiles()
    Dim doc As Document
    Dim fld As Field
    For Each doc In Application.Documents
        For Each fld In doc.Fields
            fld.Update
        Next fld
    Next doc
End Sub