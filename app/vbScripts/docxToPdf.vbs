Dim objWord, objDoc
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' Get input and output file paths from command-line arguments
inputFile = WScript.Arguments(0)
outputFile = WScript.Arguments(1)

' Open the document
Set objDoc = objWord.Documents.Open(inputFile)

' Save as PDF (format 17)
objDoc.SaveAs outputFile, 17

' Close document and quit Word
objDoc.Close False
objWord.Quit