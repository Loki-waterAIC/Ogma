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

' Run the macro
objWord.Run macroName

' Save and close the document
objDoc.Save
objDoc.Close

' Quit Word
objWord.Quit

If Err.Number <> 0 Then
    WScript.Echo "Error closing Word: " & Err.Description
    Err.Clear
End If

' Clean up
Set objDoc = Nothing
Set objWord = Nothing
