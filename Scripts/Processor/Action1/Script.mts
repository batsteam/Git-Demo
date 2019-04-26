





'***************************************************************************************************************************************************************************
' @@ hightlight id_;_Browser("Browser").Page("Case - Consultations").WebElement("Queue Name")_;_script infofile_;_ZIP::ssf15.xml_;_
	'SCRIPT NAME:  Processor
	'DESCRIPTION: 	This Script Dynamically loads Repositories and Functions based on Procedures in the running Test Case spreadsheet, 
	'				then Exports Run Results to the Test Case Results spreadsheet																
	'Last Updated: 8/10/2018				
	'Updated by: Geoff Viado / gviado@humana.com					
'***************************************************************************************************************************************************************************
'SystemUtil.CloseProcessByName "excel.exe"						'Close all excel processes

'Initialize TC variables
Environment.Value("TestCase") = DataTable("TestCase",dtGlobalSheet)
Environment.Value("Environment") = LCase(DataTable("Environment",dtGlobalSheet))
Environment.Value("Application") = LCase(DataTable("Application",dtGlobalSheet))
Environment.Value("Browser") = LCase(DataTable("Browser",dtGlobalSheet))
Environment.Value("EmailResults") = LCase(DataTable("EmailResults",dtGlobalSheet))

'Determine Action Sheet
Environment.Value("SheetCount") = DataTable.GetSheetCount	
	If Environment.Value("SheetCount") = 2 Then
		Environment.Value("ActionSheet") = "Processor"
	Elseif Environment.Value("SheetCount") = 3 Then
		Environment.Value("TCSheet")=DataTable.GetSheet(3).Name
			If Environment.Value("TCSheet") = "Processor [Processor]" Then
				Environment.Value("ActionSheet") = "Processor [Processor]"
			Else
				DataTable.DeleteSheet 3	
				Environment.Value("ActionSheet") = "Processor"
			End If	
	End If
	
'Path Definitions 
  Set objFSO = CreateObject("Scripting.FileSystemObject")
	Environment.Value("MainPath")= objFSO.GetParentFolderName(objFSO.GetParentFolderName(Environment.Value("TestDir")))
  Set objFSO=Nothing
  
Environment.Value("ScriptPath") = Environment.Value("MainPath") & "\Scripts\"		'Define location of scripts
Environment.Value("RepositoryPath") = Environment.Value("MainPath") & "\ObjectRepository\"		'Define location of object repository
Environment.Value("TestSet") = "\TestCaseFiles\"		'Define Test Set folder (location of test cases and suites)
Environment.Value("TestPath") = Environment.Value("MainPath")&Environment.Value("TestSet")&Environment.Value("TestCase")		'Define Test Case folder
Environment.Value("StepResultsPath") = Environment.Value("TestPath")&"\Results\" & Environment.Value("TestCase") & "_StepResults.xls"	'Define StepResults path
Environment.Value("TestResultsPath") = Environment.Value("TestPath")&"\Results\" & Environment.Value("TestCase") & "_TestResults.xls"	'Define TestResults path
Environment.Value("FilePath") = Environment.Value("TestPath")&"\"&Environment.Value("TestCase")&".xlsx"	'Define File path

'Load Common Files - Config, CommonFunctions, GlobalVariables
LoadFunctionLibrary Environment.Value("ScriptPath")&"Config.qfl",Environment.Value("ScriptPath")&"CommonFunctions.qfl",Environment.Value("ScriptPath")&"GlobalVariables.qfl"

'Call BeforeTest	'Setup test environment

	importAnyXL Environment.Value("FilePath")		'Import test cases from XL to DataTable
	Environment.Value("GLconstNumProcedures") = DataTable.GetSheet(Environment.Value("TestCase")).GetRowCount	'Get number of test case steps
	
	'Initialize Result Row for Results Exporting  
	Environment.Value("ResultRow") = 1
	
	print "----- Test Case: " & Environment.Value("TestCase")
	
	'For Each Step in Test Case Spreadsheet, Import Needed Functions and Repository
	On Error Resume Next
	For procedureCount = 1 to Environment.Value("GLconstNumProcedures")
	Environment.Value("procedureCount") = procedureCount
	
	    If DataTable("Module",Environment.Value("TestCase")) <> "" AND UCase(DataTable("Skip",Environment.Value("TestCase"))) <> "Y" Then	
			Environment.Value("GLconstProcedureName") = DataTable("Module",Environment.Value("TestCase"))	'Module
			Environment.Value("GLvarProcedureName") = DataTable("Procedure",Environment.Value("TestCase"))	'Procedure
			Environment.Value("skipStep") = DataTable("Skip",Environment.Value("TestCase"))		'Skip
			GLvarTestStep = DataTable("Case",Environment.Value("TestCase"))		'Case
				Environment.Value("GLvarTestStep") = GLvarTestStep
			GLvarTestData = DataTable("Scenario",Environment.Value("TestCase"))				'Scenario
				Environment.Value("GLvarTestData") = GLvarTestData
			Environment.Value("strTest") = "Test Step "&Environment.Value("procedureCount")&"  "&Environment.Value("GLvarProcedureName")&" "&Environment.Value("GLvarTestStep")&" "&DataTable("Scenario",Environment.Value("TestCase"))
					
			'Load Function
			LoadFunctionLibrary Environment.Value("ScriptPath")&DataTable("Module",Environment.Value("TestCase"))&"\"&Trim(DataTable("Procedure",Environment.Value("TestCase")))&".qfl"
	
				If RepositoriesCollection.Count > 0 Then    'If OR is loaded
					If StrComp(RepositoriesCollection.Item(1), Environment.Value("RepositoryPath") &Environment.Value("GLconstProcedureName")&"\"&Environment.Value("GLvarProcedureName")&".tsr", 1) <> 0 Then   'If loaded OR and new OR are not equal
						RepositoriesCollection.Remove(1)					'Release loaded Object Respository
						RepositoriesCollection.Add(Environment.Value("RepositoryPath") &Environment.Value("GLconstProcedureName")&"\"&Environment.Value("GLvarProcedureName")&".tsr")   'Load corresponding object repository
					End If
				
				Else	'No OR loaded so load OR
					RepositoriesCollection.Add(Environment.Value("RepositoryPath") &Environment.Value("GLconstProcedureName")&"\"&Environment.Value("GLvarProcedureName")&".tsr")   ' Load corresponding object repository
				End If
				
				print "Step " & procedureCount & "/" & Environment.Value("GLconstNumProcedures") & "  "&Environment.Value("GLvarProcedureName")&" "&Environment.Value("GLvarTestStep")&" "&DataTable("Scenario",Environment.Value("TestCase"))	'Display step in print log
				Execute Environment.Value("GLvarProcedureName")		'Run procedure within qfl file					
				Call WriteResults 				'Output Results into the Processor datasheet
				
				Environment.Value("ResultRow") = Environment.Value("ResultRow") + 1	'Write results in next row	
				DataTable.GetSheet(Environment.Value("TestCase")).SetNextRow 	'Execute next row of the test case	
	
					If Err.number <> 0 then 				
						Reporter.ReportEvent micWarning,"Error Occured","Error number is  " &  err.number & " and description is : " &  err.description   
						err.clear
					End If
	
		ElseIf  DataTable("Module",Environment.Value("TestCase")) <> "" AND UCase(DataTable("Skip",Environment.Value("TestCase"))) = "Y" Then		'Skip and exec next row
				DataTable.GetSheet(Environment.Value("TestCase")).SetNextRow 	'Execute next row of the test case
	
		Else	'Error in test case file
				ReportEvent micFail, "INPUT FILE ERROR", "PLEASE SPECIFY THE MODULE. TEST STEP:  " & procedureCount
				Exit For
		End If
	Next

Call AfterTest 'Export reports
