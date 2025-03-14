
Sub UpdateAllFields()
    Dim doc As Document
    Set doc = ActiveDocument

    With ActiveDocument

        ' Update fields in the main document body
        .Fields.Update

        ' Update fields in headers And footers For each section
        Dim oSection            As Section
        Dim oHeader             As HeaderFooter
        Dim oFooter             As HeaderFooter
        Dim oBOK_ID             AS "BOK ID"
        Dim oDocument_Name      AS "Document Name"
        Dim oCompany_Name       AS "Company Name"
        Dim oDivision           AS "Division"
        Dim oAuthor             AS "Author"
        Dim oCompany_Address    AS "Company Address"
        Dim oProject_Name       AS "Project Name"
        Dim oProject_Number     AS "Project Number"
        Dim oEnd_Customer       AS "End Customer"
        Dim oSite_Name          AS "Site Name"
        Dim oFile_Name          AS "File Name"
        For Each oSection In .Sections
            oSection.Range.Fields.Update
            For Each oHeader In oSection.Headers
                oHeader.Range.Fields.Update
            Next oHeader
            For Each oFooter In oSection.Footers
                oFooter.Range.Fields.Update
            Next oFooter
            FOR Each 
        Next oSection
    End With
End Sub
