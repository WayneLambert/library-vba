'Kill all files of (or not of) a certain Filetype
Sub KillFiles()
'Delete all but *.png files in the HTML folder
'To kill all *.png files, then replace <> with =
gFile = Dir(DirName & "\*.*")
Do While gFile <> ""
	If Right(gFile, 3) <> "png" Then Kill DirName & "\" & gFile
	gFile = Dir
Loop
End Sub