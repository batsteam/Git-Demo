'-----------------------------------------------
'Author: Geoff Viado
'Update Date: 11/4/14	by Geoff Viado
'UFT Runner
'-----------------------------------------------

Dim filePath, fileName, qtFile
Set qtApp=CreateObject("QuickTest.Application")
	
qtApp.Launch
qtApp.Visible = True		'True = UFT visible, False = UFT silent mode
	
	fileName = "Processor"	'Set processor name
	filePath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)-Len(WScript.ScriptName)))
	'WScript.Echo filePath

	qtFile = filePath&fileName
	'WScript.Echo qtFile
	
qtApp.Open qtFile, False	'Open test, True = file read-only, False = file editable

Set qtTest=qtApp.Test

qtTest.Run		'Execute test
WScript.Sleep 12000
'qtTest.Close	'Close test file
'qtApp.Quit		'Close UFT
Set qtTest=Nothing
Set qtApp=Nothing