'Loading Environment Variables in runtime
Environment.LoadFromFile "C:\UFT_Automation\Config\SFDC_Global_Variables.xml"

'Closing Browser if any open
SystemUtil.CloseProcessByName Environment("ChromeBrowser")

'Closing Excel files if any open
SystemUtil.CloseProcessByName "Excel.exe"

'Loading Function Libraries and Object Repeositry in runtime
LoadFunctionLibrary Environment("Application_Functions_Service")
LoadFunctionLibrary Environment ("Generic_Functions_Service")
'RepositoriesCollection.Add Environment ("Service_Object_Repository")

SystemUtil.Run Environment("ChromeBrowser"),Environment("Test_URL"),"","open",3
'SystemUtil.Run Environment("ChromeBrowser"),Environment("Prod_URL"),"","open",3

'Login to SF Classic as an Admin profile
LoginToSFDC_Admin 

'Opening Control file
Set oXl=createobject("excel.application")
'oXl.Workbooks.Open Environment ("Control_File_Service_Cert")
'oXl.Workbooks.Open Environment ("Control_File_Service_Cert2")
'oXl.Workbooks.Open Environment ("Control_File_Service_Cert4")
oXl.Workbooks.Open Environment ("Control_File_Service_Prod")
'oXl.Workbooks.Open Environment ("Control_File_Service_ProdTest")

oXl.visible=true

Set oTC=oXl.ActiveWorkbook.Worksheets(1)
Set oTS=oXl.ActiveWorkbook.Worksheets(2)

'Finding size of the rows in "Test Case" sheet and "Test Steps" sheet
sizeTC=oTC.usedRange.rows.count
sizeTS=oTS.usedRange.rows.count

''Clearing content of Pass/Fail status and date/time columns of "Test Case" sheet
'For TCi = 2 To sizeTC
'	oTC.cells(TCi,6) = ""
'	oTC.cells(TCi,7) = ""
'	oTC.cells(TCi,6).interior.colorindex = 0
'Next
'
''Clearing content of Pass/Fail status and date/time columns of 'Test Steps' sheet
'For TSi = 2 To sizeTS
'	oTS.cells(TSi,13) = ""
'	oTS.cells(TSi,14) = ""
'	oTS.cells(TSi,13).interior.colorindex = 0
'Next
'
'Looping through 'Test Case' sheet of Control file
	For rowTC = 2 To 12 'sizeTC
		'Assigning 'Execute Flag' column of the Test Case sheet
		executeFlag=oTC.cells(rowTC,1)
		
		'Assigning 'Pass' to 'status_TC' variable
		status_TC="Pass"
		
		'If Execute Flag 'YES', then continue with the test
		If ucase(executeFlag)="Y" Then
			'Assigning 'User Story ID' and 'Test Case ID' column of the Test Case sheet
			USID_TC=oTC.cells(rowTC,3)
			TCID_TC=oTC.cells(rowTC,4)
			
			'Looping through 'Test Steps' sheet of Control file
			For rowTS = 2 To sizeTS
				'Assigning 'User Story ID' and 'Test Case ID' column of the Test Steps sheet
				USID_TS=oTS.cells(rowTS,1)
				TCID_TS=oTS.cells(rowTS,2)
				'Continuing the test only if 'User Story ID' and 'Test Case ID' column of the Test Case sheet equals to 
											 'User Story ID' and 'Test Case ID' column of the Test Steps sheet
				If USID_TC=USID_TS and TCID_TC=TCID_TS Then
					'Assigning Keyword and Arguments columns to the variables
					KeyWord=oTS.cells(rowTS,4)
					arg1=oTS.cells(rowTS,5)
					arg2=oTS.cells(rowTS,6)
					arg3=oTS.cells(rowTS,7)
					arg4=oTS.cells(rowTS,8)
					arg5=oTS.cells(rowTS,9)
					arg6=oTS.cells(rowTS,10)
					arg6=oTS.cells(rowTS,11)
					arg6=oTS.cells(rowTS,12)
					
					'KeyWordExecutor - executes corresponding functions and dynamically returns Pass or Fail for each executed function				
					TestStepStatus = KeyWordExecutor(KeyWord,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8)
					
					'After each execution, Pass or Fail will be written to 'Status' column of the 'Test Steps' sheet
					oTS.cells(rowTS,13) = TestStepStatus
					
					'After each execution, based on Pass or Fail status, it will insert color green or red color index
					if oTS.cells(rowTS,13) = "Pass" then
						oTS.cells(rowTS,13).interior.colorindex=4
					Else
						oTS.cells(rowTS,13).interior.colorindex=3
					End If
					
					'After each execution, time will be written to 'Date/Time' column of the 'Test Steps' sheet
						oTS.cells(rowTS,14)=now
						
					'If any step in the execution fails, the 'Fail' value will be assigned to 'status_TC' variable
					If TestStepStatus = "Fail" Then				
						status_TC = "Fail"
					End If
				
				End If			
			Next
			
				'If 'status_TC' variable contains 'Pass' or 'Fail' value, 'Pass' or 'Fail' value will be written to 'Status' column of the 'Test Case' sheet
				oTC.cells(rowTC,6)=status_TC
				oTC.cells(rowTC,7)=now
				If lcase(status_TC)="pass" Then
				oTC.cells(rowTC,6).interior.colorindex=4
				else
				oTC.cells(rowTC,6).interior.colorindex=3
				End If
		else
			'If 'Test Case ID' column of the Test Case sheet does not equal to 'Test Step ID' column of the Test Steps sheet, 
			'Not Executed' value will be assigned to 'STATUS' variable
			oTC.cells(rowTC,6)="Not Executed"	
		End If
	Next

'	'Saving Control file with test results in 'Test_Result_Excel' folder
oXl.ActiveWorkbook.Save
'oXl.ActiveWorkbook.SaveAs Environment("excelResultPath") & "SF_SERVICE_REGRESSION_Test_Results_" & timeStamp & ".xlsx"

'	'Closing Control File
oXl.Workbooks.Close
oXl.Quit
''
''Removing reference from the objects
Set oTC=nothing
Set oTS=nothing
Set oXl=nothing

'==============================================


