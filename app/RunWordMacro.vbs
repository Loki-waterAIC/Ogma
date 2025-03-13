Dim objWord, objDoc, macroName, docPath

' Get arguments from command line
Set objArgs = WScript.Arguments
If objArgs.Count <> 2 Then
    WScript.Echo "Usage: RunWordMacro.vbs <docPath> <macroName>"
    WScript.Quit 1
End If

docPath = objArgs(0)
macroName = objArgs(1)

' Create Word Application object
Set objWord = CreateObject("Word.Application")
objWord.Visible = False  ' Set to True if you want to see Word opening

' Open the Word document
Set objDoc = objWord.Documents.Open(docPath)

' Run the macro
objWord.Run macroName

' Save and close the document
objDoc.Save
objDoc.Close

' Quit Word
objWord.Quit

' Clean up
Set objDoc = Nothing
Set objWord = Nothing
