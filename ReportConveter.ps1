#powershell script sits inside the same path as the report.accdb you want to convert
#Removes the old report.mdb file so you can keep reusing the script for new accdb files, if not needed remove line 3 completely
Remove-Item "$psscriptroot\report.mdb"
$Access = New-Object -com Access.Application
$Access.ConvertAccessProject(
		#accdb file to convert
		"$psscriptroot\report.accdb",
		#converted file name
		"$psscriptroot\report.mdb",
		"acFileFormatAccess2000")
$Access.Quit()
