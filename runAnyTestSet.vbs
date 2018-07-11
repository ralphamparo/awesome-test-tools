'Author: Ralph Amparo
'Description: Run multiple test sets using OTA library. This library will rerun any scripts that encountered server connectivity issues during run time.

'declare global constants
Const qcHostName = ""
Const qcServer = "http:// /qcbin/"'ALM server name
Const qcDomain = ""'ALM domain name
Const qcProject = " "'ALM project name
Const qcUser = "" 'ALM user name
Const qcPassword = "" 'ALM password

'folder path to get the test name(s) for execution
Const testFolderPath = "Root\ \"
'test set name array, modify this to execute any test set(s)
testSetNameArray = Array("ATC2_Config_Rap","ATC1_Status_Rap")

Const testCaseIDVal="" 'set this to blank for bulk execution
runAnyTestSet(testCaseIDVal) ' run all the test sets define in testPathArray 

'function for running any test sets
'Parameters: otdc-OTA object, tsFolderName - folder path, tSetName - test set to be executed, testCaseID - test ID of the case for execution
Public Sub RunTestSet(otdc, tsFolderName, tSetName,testCaseID)
	
	Dim TSetFact, tsList
	Dim theTestSet
	Dim tsTreeMgr
	Dim tsFolder
	Dim Scheduler
	Dim nPath
	Dim execStatus
	
'Get the test set tree manager from the test set factory
'tdc is the global TDConnection object.
	Set TSetFact = otdc.TestSetFactory
	Set tsTreeMgr = otdc.TestSetTreeManager
' Get the test set folder passed as an argument to the example code
'nPath = "Root \ " & Trim(tsFolderName)
	nPath = Trim(tsFolderName)
	Set tsFolder = tsTreeMgr.NodeByPath(nPath)
	If tsFolder Is Nothing Then
		Err.Raise vbObjectError + 1, "RunTestSet", "Could not find folder " & nPath
	End If
	
' Search for the test set passed as an argument to the example code
	
	Set tsList = tsFolder.FindTestSets(tSetName)
'error handling for invalid test set names
	If tsList Is Nothing Then
		Err.Raise vbObjectError + 1, "RunTestSet", "Could not find test set in the " & nPath
		Exit Sub
	End If
'error handling for multiple test names being set
	If tsList.Count > 1 Then
		MsgBox "FindTestSets found more than one test set: refine search"
		Exit Sub
	ElseIf tsList.Count < 1 Then
		MsgBox "FindTestSets: test set not found"
		Exit Sub
	End If
	
'set TestSet object to get the test set ID
	Set theTestSet = tsList.Item(1)
	
'Start the scheduler on the local machine
	Set Scheduler = theTestSet.StartExecution("")
'Run all tests on the local machine
	Scheduler.RunAllLocally = True
'if testCaseID parameter is "" run the whole test set
	If testCaseID = "" Then
'Run the tests
		Scheduler.Run
	ElseIf testCaseID <> "" Then
'run all not completed scripts, Scheduler.Run(CStr(TC_TESTCYCL_ID)) ..
		Scheduler.Run(Cstr(testCaseID))
	End If
	
'get the execution status of the test set
	Set execStatus = Scheduler.ExecutionStatus
'update the status while the execution is not finished
	While (RunFinished = False)
		execStatus.RefreshExecStatusInfo "all", True
		RunFinished = execStatus.Finished
'System.Threading.Thread.Sleep(10000)
	Wend
'update the status again after execution ends
	execStatus.RefreshExecStatusInfo "all", True
	
'determine which scripts are in "Not Completed" or in "No Run" status.
	Set TestSetF = tsFolder.TestSetFactory
	Set TestSetL = TestSetF.NewList("")
'traverse through each test case in the test set, get its execution status after the first batch run
	For Each TestSetObj In TestSetL
		
		TestSetObj.Refresh'refresh object to make sure its properties are changed before validation
		
		Set testSetFilter = otdc.TestSetFactory.Filter
		testSetFilter.Filter("CY_CYCLE_ID") = "'" & theTestSet.ID & "'"
		Set testSetList = testSetFilter.NewList
		
		Set TestCaseF = testSetList(1).TSTestFactory
		Set TestCaseL = TestCaseF.NewList("")
		
		For Each TestCaseObj In TestCaseL
			testCasePassed = Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Passed")
			testCaseFailed = Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Failed")
			testCaseExecutionDate = CDate(TestCaseObj.LastRun.Field("RN_EXECUTION_DATE"))
'testNotYetExecuted = DateDiff("d",testCaseExecutionDate,Date())
			
'if test status wasn't passed or failed. rerun the script by getting the testobjID, do recursion 
			If Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Passed") <= 0 And Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Failed",1) <= 0 Then
				testCaseIdValue=TestCaseObj.ID 'get test case ID of the test case in "No Run" or "Not Completed Status"
				Exit For
			End If 
		Next
		Exit For ' exit loop for the first instance of the TestSetObj
	Next
	
'disconnectTDObject(otdc)
	killALMPRocesses
'Sleep(10000)' wait for 10 seconds
' if a test case is in Not completed or in No Run Status or if the script encountered an error do recursion
	if testCaseIDValue <> "" Or Err.Number > 0 Then
		runAnyTestSet Cstr(testCaseIDValue)
	Elseif testCaseIDValue = "" and testCaseId <> "" And Err.Number > 0 Then
		runAnyTestSet Cstr(testCaseID)
	End If
	
End Sub


'function for error handling when the connection to the server has failed during run time. This will run all the scripts that have not finished execution.
'Parameters: tSetName - test set to be executed, testCaseID - test ID of the case for execution
Function getNotCompletedTestCaseId(tSetName,testCaseID)
	
	runAnyTestSet(testCaseID)
	Set otdc = CreateObject("tdapiole80.tdconnection") ' set tdc object
	
	otdc.InitConnectionEx qcServer
	otdc.Login qcUser, qcPassword
	otdc.Connect qcDomain, qcProject
	
	Set TSetFact = otdc.TestSetFactory
	Set tsTreeMgr = otdc.TestSetTreeManager
	Set tsFolder = tsTreeMgr.NodeByPath(testFolderPath)
	Set TestSetF = tsFolder.TestSetFactory
	Set TestSetL = TestSetF.NewList("")
	
	Set tsList = tsFolder.FindTestSets(tSetName)
	Set theTestSet = tsList.Item(1)
	testSetID = Cstr(theTestSet.ID)
	
	For Each TestSetObj In TestSetL
		
		TestSetObj.Refresh
		
		Set testSetFilter = otdc.TestSetFactory.Filter
		testSetFilter.Filter("CY_CYCLE_ID") = "'" & theTestSet.ID & "'"
		Set testSetList = testSetFilter.NewList
		
		Set TestCaseF = testSetList(1).TSTestFactory
		Set TestCaseL = TestCaseF.NewList("")
		
		For Each TestCaseObj In TestCaseL
'Msgbox TestCaseObj.ID
			testCasePassed = Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Passed")
			testCaseFailed = Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Failed",1)
			testCaseExecutionDate = CDate(TestCaseObj.LastRun.Field("RN_EXECUTION_DATE"))
			testNotYetExecuted = DateDiff("d",testCaseExecutionDate,Date())
'if test status wasn't passed or failed. rerun the script by getting the testobjID, do recursion
			
			If Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Passed") <= 0 And  Instr(1,TestCaseObj.LastRun.Field("RN_STATUS"),"Failed",1) <= 0 and CLNG(TestCaseObj.ID) <> Clng(testCaseID) Then
					testCaseIDValue=Cstr(TestCaseObj.ID)
					Exit For
			End If 
		Next
		Exit For
	Next
	
	if testCaseIDValue <> "" Then'if a test case with "Not Completed" or "No Run" exists. Do recursion
'killALMPRocesses
			getNotCompletedTestCaseId tSetName,testCaseIDValue
'RunTestSet tdc, testFolderPath, testSetNameArray(testSetNameCtr),testCaseIDValue
	End If
End Function

'main function for running the test sets
Sub runAnyTestSet(testCaseID)
	killALMPRocesses()'kill ALM processes for reruns
	
	Set tdc = CreateObject("tdapiole80.tdconnection") ' set tdc object
	If (tdc Is Nothing) Then
		MsgBox "tdc object is empty"
		Exit Sub
	End If
	
	tdc.InitConnectionEx qcServer'connect to the ALMs server
	
	If Err.Number > 0 Then
		For testSetNameCtr = 0 To UBound(testSetNameArray) 
			getNotCompletedTestCaseId testSetNameArray(testSetNameCtr),""
		Next 
'Exit Sub
	End If
	
	tdc.Login qcUser, qcPassword ' login function
	tdc.Connect qcDomain, qcProject ' connect to domain and project function
	
'run all test sets specified in the testSetNameArray
	For testSetNameCtr = 0 To UBound(testSetNameArray) 
		RunTestSet tdc, testFolderPath, testSetNameArray (testSetNameCtr),testCaseID
	Next
	
	disconnectTDObject(tdc)' disconnect from ALM
	killALMPRocesses
End Sub

Sub disconnectTDObject(tdc)
'Disconnect from the project
	If tdc.Connected Then
		tdc.Disconnect
	End If
'Log off the server
	If tdc.LoggedIn Then
		tdc.Logout
	End If
'Release the TDConnection object.
	tdc.ReleaseConnection
'"Check status (For illustrative purposes.)"
	Set tdc = Nothing
End Sub

'kill all processes that are related to ALM and QTP
Sub killALMPRocesses()
	Dim objWMIService, objProcess, colProcess 
	Dim strComputer, strProcessKill 
	strComputer = "." 
	
	strProcessArray=Array("'bp_exec_agent.exe'","'wexectrl.exe'","'QTPro.exe'","'AQTRmtAgent.exe'")
'strProcessKill = "'calc.exe'" 
	
	For processIndex = 0 to Ubound(strProcessArray)
'TerminateEXE(strProcessArray(processIndex))
		Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
		Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessArray(processIndex))
		For Each objProcess in colProcess 
			objProcess.Terminate() 
		Next 
	Next
	
End Sub