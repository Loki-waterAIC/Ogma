If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript CheckSpecificMacro.vbs <MacroName>"
    WScript.Quit -1
End If

macroName = WScript.Arguments(0)

Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' Check the Word document
Set objDoc = objWord.Documents.Open("C:\path\to\your\document.docx")
CheckMacro objDoc, macroName
objDoc.Close False

' Check the normal.dotm template
Set objTemplate = objWord.NormalTemplate
CheckMacro objTemplate, macroName

objWord.Quit

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