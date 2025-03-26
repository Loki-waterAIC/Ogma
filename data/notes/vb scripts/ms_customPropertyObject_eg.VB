' https://learn.microsoft.com/en-us/office/vba/api/word.customproperties

Sub SmartTagsProps() 
    Dim docNew As Document 
    Dim stgTag As SmartTag 
    Dim stgProp As CustomProperty 
    Dim intTag As Integer 
    Dim intProp As Integer 

    Set docNew = Documents.Add 

    'Create heading info in New document 
    With docNew.Content 
        .InsertAfter "Name" & vbTab & "Value" 
        .InsertParagraphAfter 
    End With 

    'Loop through smart tags in current document 
    For intTag = 1 To ActiveDocument.SmartTags.Count 

        With ActiveDocument.SmartTags(intTag) 

            'Verify that the custom properties 
            'For smart tags is greater than zero 
            If .Properties.Count > 0 Then 

                'Loop through the custom properties 
                For intProp = 1 To .Properties.Count 

                    'Add custom Property name To New document 
                    docNew.Content.InsertAfter .Properties(intProp) _ 
                    .Name & vbTab & .Properties(intProp).Value 
                    docNew.Content.InsertParagraphAfter 
                Next 
            Else 

                'Display message If there are no custom properties 
                MsgBox "There are no custom properties For the " & _ 
                "smart tags in your document." 
            End If 
        End With 
    Next 

    'Convert the content in the New document into a table 
    docNew.Content.Select 
    Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=2 

End Sub