' Document property struct
Type docProp
    name As String
    defaultVal As String
End Type

' Check if document property exists in current document
Function propertyExists(propName) As Boolean

    Dim tempObj
    On Error Resume Next
    Set tempObj = ActiveDocument.CustomDocumentProperties.Item(propName)
    propertyExists = (Err = 0)
    On Error GoTo 0

End Function

' Main function
Sub UpdateDocProperties(control As IRibbonControl)

' Initialize array of properites and default values
Dim props(11) As docProp

' Populate array with props and default vals
props(0).name = "BOK ID"
props(0).defaultVal = "WMLSI.XX.XX.XXX.X"

props(1).name = "Document Name"
props(1).defaultVal = "Document Name"

props(2).name = "Company Name"
props(2).defaultVal = "W. M. Lyles Co."

props(3).name = "Division"
props(3).defaultVal = "System Integration Division"

props(4).name = "Author"
props(4).defaultVal = "Lastname, Firstname"

props(5).name = "Company Address"
props(5).defaultVal = "9332 Tech Center Drive, Suite 200 | Sacramento, CA 95826"

props(6).name = "Project Name"
props(6).defaultVal = "Project Name"

props(7).name = "Project Number"
props(7).defaultVal = "WMLSI.XX.XX.XXX.X"

props(8).name = "End Customer"
props(8).defaultVal = "End Customer"

props(9).name = "Site Name"
props(9).defaultVal = "Site Name"

props(10).name = "File Name"
props(10).defaultVal = "DocumentFileName"

' Loop through array and check if each prop exists, add it if it doesn't
For i = 0 To 10
    If Not propertyExists(props(i).name) Then
        With ActiveDocument.CustomDocumentProperties
            .Add name:=props(i).name, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=props(i).defaultVal
        End With
    End If
Next i
            

' Show form
DocPropsForm.Show
End Sub
