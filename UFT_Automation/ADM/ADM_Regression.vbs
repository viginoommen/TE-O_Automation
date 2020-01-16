Dim uftAPP
'Create QTP/UFT Object
Set uftAPP = CreateObject("Quicktest.Application")
'Launch UFT and Make it Visible
    uftAPP.Launch
    uftAPP.Visible = True
	'Open the Test
    uftAPP.Open "C:\Users\sg0303193\Documents\UFT_AUTOMATION\Test_Scripts\SF_SALES_REGRESSION"
    'Run the Test
    uftAPP.Test.Run 
    'Close the Test
    uftAPP.Test.Close
    'Close the UFT tool
    uftAPP.Quit