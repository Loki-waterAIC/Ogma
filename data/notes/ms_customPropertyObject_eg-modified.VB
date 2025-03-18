'base off the examples found at https://learn.microsoft.com/en-us/office/vba/api/word.customproperties

Sub ReturnPropsAndValues()
    Dim docNew As Document
    Dim stgTag As SmartTag
    Dim stgProp As CustomProperty
    Dim intTag As Integer
    Dim intProp As Integer

    'Loop through smart tags in current document
    For intTag = 1 To ActiveDocument.SmartTags.Count

        With ActiveDocument.SmartTags(intTag)

            'Verify that the custom properties
            'For smart tags is greater than zero
            If .Properties.Count > 0 Then

                'Loop through the custom properties
                For intProp = 1 To .Properties.Count

                    With ActiveDocument.SmartTags(intTag).Properties(intProp) 
                        MsgBox "The Smart Tag name is: " & .Name & vbLf & .Value 
                    End With 
                Next
            Else

                'Display message If there are no custom properties
                MsgBox "There are no custom properties For the " & _
                "smart tags in your document."
            End If
        End With
    Next
End Sub