Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WSCript.shell")
sFolder = GetFolderPath()

If MsgBox("Archive (rar) all bak files in " & sFolder  & " folder?", vbYesNo + vbQuestion) = vbYes Then
	ZipArchive sFolder
	MsgBox "Done!"
End If


Sub ZipArchive(sFolder)
  Set oFolder = fso.GetFolder(sFolder)
  For each oFile in oFolder.Files
    If Right(oFile.Name,4) = ".bak" Then 
	oShell.run """C:\Program Files\WinRAR\WinRAR.exe"" a " & Replace(oFile.Path,".bak",".rar") & " " & oFile.Path, 0 , True
	oFile.Delete
    End If
  Next
End Sub

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder
End Function