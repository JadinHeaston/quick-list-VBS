'This script will search a given folder and return a text file of file names (plus extensions).
'The list text document is left within the directory of the script.

'Don't be udmb.
option explicit

'Declaring stupid dumb variable because VBS is a dumb baby language.
dim stupidVariable

'File system handles and variables.
dim objFSO, objFile, objFolder
Set objFSO = CreateObject("Scripting.FileSystemObject")

'This is the basename of a file. Only the file name, no extension.
dim objFileName

'Defining list text file variable and creating the text file.
dim list
Set list = objFSO.OpenTextFile("LIST.txt", 2, True, 0)

'Creating handle to browse folders.
dim objShell
Set objShell = CreateObject("Shell.Application")

'Creating message box to notify user what the script does and the general "how it works".
stupidVariable = MsgBox("This script will create a text file list of all file names (and extensions) within a folder. (This EXCLUDES folders and their names. Only files are shown.)", 1, "NOTICE")
If stupidVariable = 1 Then
	'Creating browser 
	Set objFolder = objShell.BrowseForFolder(0, "What folder would you like to create a list of?", 1, 0)


	If Not (objFolder Is Nothing) Then
		'Making this a proper collection.
		Set objFolder = objFSO.GetFolder(objFolder.Self.path + "\").Files
		'For every file in the folder, get the basename and put that in a new line in the "List.txt" file.
		For Each objFile in ObjFolder
			objFileName = objFSO.GetBaseName(objFile)
			list.WriteLine(objFileName & "." & objFSO.GetExtensionName(objFile))
		Next
	Else
		stupidVariable = MsgBox("No input given!", 0, "Quick List Failed!")
		WScript.Quit()
	End If

	'Announcing that everything is done!
	objFile = objFSO.GetFile("LIST.txt")
	MsgBox("Finished! The text file is located at: " & objFile)
Else
	stupidVariable = MsgBox("Script aborted!", 0, "Aborted!")
	WScript.Quit()
End If

'Closing the text file and the script.
list.Close
Wscript.quit()