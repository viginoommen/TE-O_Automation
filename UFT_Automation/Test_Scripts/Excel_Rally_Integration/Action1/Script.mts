SystemUtil.Run Environment("ChromeBrowser"),Environment("URL_Rally"),"","open",3
wait 10
SystemUtil.CloseProcessByName "Excel.exe"
Set XL = CreateObject("Excel.Application")
XL.Workbooks.Open "C:\UFT_Automation\Regression_suit\RallyExcelIntegration.xlsx"              
Set WSH = XL.ActiveWorkbook.Worksheets(1)
XL.Visible = True
RowCount = WSH.UsedRange.Rows.Count

For i = 1 To RowCount

	Defect_Number		= WSH.Cells(i,1)
	Name_Target			= WSH.Cells(i,2)

Find_WebEdit "smb-TextInput-input", 0, "True", "set_value", Defect_Number, "Search..."
wait 2
Find_WebEdit "smb-TextInput-input", 0, "True", "click", "", "Search..."
wait 2
Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
	myDeviceReplay.PressKey 28	
Focus_WebElement "", "Target Build", 0
wait 2
Browser("title:=.*").Page("title:=.*").WebElement("class:=smb-TextInput-renderedText", "html tag:=DIV", "visible:=True", "index:=3").click
wait 2
'to clear the field
'Browser("title:=.*").Page("title:=.*").WebEdit("class:=smb-TextInput-input", "html tag:=INPUT", "visible:=True", "index:=7").set ""
'Browser("title:=.*").Page("title:=.*").WebElement("class:=smb-TextInput-renderedText", "html tag:=DIV", "visible:=True", "index:=3").click
'wait 2
Browser("title:=.*").Page("title:=.*").WebEdit("class:=smb-TextInput-input", "html tag:=INPUT", "visible:=True", "index:=7").set Name_Target
wait 3
If Browser("title:=.*").Page("title:=.*").WebButton("class:=smb-Button smb-Button--primary smb-Button--sm chr-QuickDetailEntityFooter-saveButton","innertext:=Save","visible:=True","index:=0").Exist(4) Then
	Find_WebButton "Save", "Save", "True", "click", 0
	wait 4
End If
Next


