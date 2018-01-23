' Export_jpg 
' exports all Grapher (*.grf) files in `work_dir` to jpg file in the same diectory. 
'
Sub Main
	Debug.Clear
	work_dir ="C:\Path\to\Grapher\Files\"

	'Create Grapher as an Object
	Dim Grapher As Object

	'Start Grapher
	Set Grapher = CreateObject("Grapher.Application")
	Grapher.Visible = True

	' get first file in dir
	file = Dir$(work_dir + "*.grf")
	While file <> ""
		' This is a dirty workaround as `object.Replace(".grf",".jpg")` does not work.
        out = Split(file, ".grf")

        out_path = work_dir + out(0) + ".jpg"
		in_path = work_dir + file
		'Open an existing template file and load worksheet
		Set _Plot = Grapher.Documents.Open(in_path)
		'DeselectAll selected objects
		_Plot.Selection.DeselectAll
		'Select all objects
		_Plot.Shapes.SelectAll
		_Plot.Export2(out_path, True,"ColorDepth=24;ColorReductionMethod=5;ExportImageFormatBGR24=1;HDPI=300;Height=1019;KeepAspect=1;KeepPixelSize=0;Quality=100;VDPI=300;Width=4227",True,"jpg")
		' close file.
		_Plot.Windows(1).Close(grfSaveChangesNo)
		file = Dir$()
	Wend
	Grapher.Quit
End Sub
