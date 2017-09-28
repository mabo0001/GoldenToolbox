Sub Main
	' Load numerous inputfiles into a Grapher template
	' save plot as *.grf file and export plots as *.jpg image 
	' 
	Debug.Clear
	template_path = "\Path\to\template\_Template.grt"
	file_path = "\Path\to\Files\"
	' list of filenames
	' linebreak in array `_` for readability
	filename_list = Array("[filename_1]", "[filename_2]", _
	..., _
	"[filename_n]" 
	)

	'Create Grapher as an Object
	Dim Grapher As Object

	'Start Grapher
	Set Grapher = CreateObject("Grapher.Application")

	Grapher.Visible = True
	For Each fname In filename_list
		' create path for each fname in list
		full_path = file_path + fname 
	
		'Open an existing template file and load worksheet
		Set _Templategrt = Grapher.Documents.Open(template_path, full_path + ".csv")

		'Save file as Grapher file with the same filename as input fil
		_Templategrt.SaveAs(full_path + ".grf")
		
		'DeselectAll selected objects
		_Templategrt.Selection.DeselectAll
		'Select all objects
		_Templategrt.Shapes.SelectAll

		'Export as image file
		_Templategrt.Export2(full_path + ".jpg",True,"ColorDepth=24;ColorReductionMethod=5;ExportImageFormatBGR24=1;HDPI=300;Height=260;KeepAspect=1;KeepPixelSize=0;Quality=100;Update_ColorDepth=24;Update_ColorReductionMethod=5;Update_Height=260;Update_Width=1092;VDPI=300;Width=1092",True,"jpg")

		
		' close file.
		_Templategrt.Windows(1).Close(grfSaveChangesNo)
	Next
End Sub
