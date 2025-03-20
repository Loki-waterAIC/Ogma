' declare variables
Dim objWord, objDoc, macroName, docPath, wordVisible

' Get arguments from command line
Set objArgs = WScript.Arguments
If objArgs.Count <> 3 Then
    WScript.Echo "Usage: RunWordMacro.vbs <docPath> <macroName> <wordVisible>"
    WScript.Quit 1
End If

docPath = objArgs(0)
macroName = objArgs(1)
wordVisible = objArgs(2)

' Create Word Application object
Set objWord = CreateObject("Word.Application")
objWord.Visible = wordVisible  ' Set to True if you want to see Word opening

' Open the Word document
Set objDoc = objWord.Documents.Open(docPath)

' check if the Macro exits
' Check the Word document
CheckMacro objDoc, macroName
objDoc.Close False

' Check the normal.dotm template
Set objTemplate = objWord.NormalTemplate
CheckMacro objTemplate, macroName

' Run the macro
objWord.Run macroName

' Save and close the document
objDoc.Save()
objDoc.Close()


' Quit Doc and Word
objDoc.Quit()
objWord.Quit()

If Err.Number <> 0 Then
    WScript.Echo "Error closing Word: " & Err.Description
    Err.Clear
End If

' Clean up
Set objDoc = Nothing
Set objWord = Nothing

WScript.Quit 0


Sub CheckMacro(obj, macroName)
    On Error Resume Next
    Set vbProj = obj.VBProject
    Set vbComp = vbProj.VBComponents(macroName)
    If vbComp Is Nothing Then
        WScript.Echo "The macro '" & macroName & "' does not exist in " & obj.Name
        WScript.Quit -1
    Else
        WScript.Echo "The macro '" & macroName & "' exists in " & obj.Name
    End If
    On Error GoTo 0
End Sub