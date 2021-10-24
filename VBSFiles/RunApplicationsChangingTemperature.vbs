'---------------------------------------------------------------------------------------------------
'	'RunApplicationsChangingTemperature.vbs'
' 	
'	Changes BVT temperature either reading from file or with specified start, end and step values.
'	On each changed temperature run the specified applications.
'
'---------------------------------------------------------------------------------------------------


'###################################################################################################
' Getting input values
'---------------------------------------------------------------------------------------------------
'Specified values
temperature_type = -1.0
to_sum = 0
starttemp = 300
tempstep = 5
endtemp = 310
waitingtime = 5
dFlow = 2000 
SerialNumber = "3280"
to_tune = 0
apps_amount = 1
ExePath    	 = "C:\Program Files\Bruker the minispec"
AppPath      = "Application path"
Dim appsfile()

'Reading input file (Comment to use specified values)
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile ("input.txt", 1)
row = 0
Do Until file.AtEndOfStream
	if row = 0 then
		temperature_type = CDbl(file.Readline)
	end if
	if row = 1 then
		starttemp = CDbl(file.Readline)
		if temperature_type = 0.0 then
			to_sum = -2
			tempstep = 1
			endtemp = starttemp
		end if
	end if
	if temperature_type = 1.0 then
		if row = 2 + to_sum then
			tempstep = CDbl(file.Readline)
		end if
		if row = 3 + to_sum then
			endtemp = CDbl(file.Readline)
		end if
	end if
	if row = 4 + to_sum then
		waitingtime = CDbl(file.Readline)
	end if
	if row = 5 + to_sum then
		dFlow = CDbl(file.Readline)
	end if
	if row = 6 + to_sum then
		SerialNumber = file.Readline
	end if
	if row = 7 + to_sum then
		to_tune = CInt(file.Readline)
	end if
	if row = 8 + to_sum then
		apps_amount = CInt(file.Readline)
		Redim appsfile(apps_amount)
	end if
	if row = 9 + to_sum then
		ExePath = file.Readline
	end if
	if row = 10 + to_sum then
		AppPath = file.Readline
	end if
	if row >= 11 + to_sum then
		ApplicFile = file.Readline
		appsfile(row - 11 - to_sum) = ApplicFile
	end if
	if row >= 10 + apps_amount + to_sum then
		exit do
	end if
	
	row = row + 1

Loop

WScript.StdOut.WriteLine( "" & temperature_type )
WScript.StdOut.WriteLine( "" & to_sum )
WScript.StdOut.WriteLine( "" & starttemp )
WScript.StdOut.WriteLine( "" & tempstep )
WScript.StdOut.WriteLine( "" & endtemp )
WScript.StdOut.WriteLine( "" & waitingtime )
WScript.StdOut.WriteLine( "" & dFlow )
WScript.StdOut.WriteLine( "" & SerialNumber )
WScript.StdOut.WriteLine( "" & to_tune )
WScript.StdOut.WriteLine( "" & apps_amount )
WScript.StdOut.WriteLine( "" & ExePath )
WScript.StdOut.WriteLine( "" & AppPath )
For Each app In appsfile
	WScript.StdOut.WriteLine( "" & app )
Next

'###################################################################################################
'The 'WinAcquisit.Embedding' Server of the WinAcquisit Software offers the methods for configuring 
'its wake up behaviour.
'---------------------------------------------------------------------------------------------------
Set uti = CreateObject( "WinAcquisit.Utilities" )
Set emb = CreateObject( "WinAcquisit.Embedding" )
emb.ShowWindow( emb.NORMAL )

'###################################################################################################
'	Declare and initialize variables
'---------------------------------------------------------------------------------------------------
DIM gblnExitOnError, gblnDontAsk
gblnExitOnError = TRUE
gblnDontAsk		= FALSE
 
'###################################################################################################
'	Get a reference to WinEPR Acquisition's BVT object
'---------------------------------------------------------------------------------------------------
Set bvt = CreateObject( "WinAcquisit.BVT" )

'###################################################################################################
'	Declare and initialize variables
'---------------------------------------------------------------------------------------------------

Dim StdIn, StdOut
Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut

'###################################################################################################
'	Set configuration
'---------------------------------------------------------------------------------------------------

'SerialNumber e Path
MinispecSerialNumberPrefix = "ND"
MinispecSerialNumber = SerialNumber
MinispecExePath    	 = ExePath

'###################################################################################################
'	Data Pairs Variable Declaration and Counter Initialization
'---------------------------------------------------------------------------------------------------
DataPairsCnt 	 = 0	' counter
DataPairsDim	 = 25 	' initial dimension
DataPairsPortion = 25	' to re dim during read in
ReDim x( dataPairsDim ), y( dataPairsDim )


'###################################################################################################
'	Get a reference to the minispec Software's PNMR object
'---------------------------------------------------------------------------------------------------
Set pnmr = CreateObject( "theMinispec.PNMR" )

'###################################################################################################
'	Header Output
'---------------------------------------------------------------------------------------------------
StdOut.WriteLine ""
StdOut.WriteLine "-----------------------------------------------------"
StdOut.WriteLine "            'RunApplicationsChangingTemperature.vbs' "
StdOut.WriteLine "-----------------------------------------------------"
StdOut.WriteLine ""
StdOut.WriteLine ""

'###################################################################################################
'	Check if BVT's On. Abort if not.
'---------------------------------------------------------------------------------------------------

StdOut.WriteLine "bvt.IsBVTOn:"
bOn = bvt.IsBVTOn
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	StdOut.WriteLine( "'bvt.IsBVTOn' ERROR #" & ErrNo & ": " & CHR(10) & ErrMsg )
	StdOut.WriteLine( "----SCRIPT ABORTED----" )
	MsgBox( "ERROR: NO ACCESS TO BVT. - Script aborted." )
	WScript.Quit
else
	if	bOn <> TRUE	then
		StdOut.WriteLine( " IsBVTOn: " & bOn )
		StdOut.WriteLine( "----SCRIPT ABORTED----" )
		MsgBox( "ERROR: BVT OFF. - Switch on, then re-start script. Script aborted." )
		WScript.Quit
	else
		StdOut.WriteLine( " IsBVTOn: " & bOn )
	end if
End If
StdOut.WriteLine ""

'###################################################################################################
'	Open access to the minispec Software's PNMR object and connect with default electronic unit
'---------------------------------------------------------------------------------------------------
'OpenPNPR - Opens the PNMR object and connects the PNMR object with the default mq
'electronic unit, if not connected already

'No arguments

'No return value
StdOut.WriteLine( "OPENING ACCESS TO THE MINISPEC SOFTWARE..." )
StdOut.WriteLine " pnmr.OpenPNMR:"
pnmr.OpenPNMR
If	pnmr.IsLastError then
	ErrMsg = pnmr.GetLastError( ErrNo )
	StdOut.WriteLine " 'pnmr.OpenPNMR' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg
else
	StdOut.WriteLine( "ACCESS TO THE MINISPEC SOFTWARE ESTABLISHED." )
End If
StdOut.WriteLine ""


'###################################################################################################
'	Connect the minispec
'---------------------------------------------------------------------------------------------------

errMsg = "Instrument Connection Failure - Program Aborted."
StdOut.WriteLine( "CONNECTING MINISPEC ELECTRONIC UNIT..." )
if	ConnectMinispec( MinispecSerialNumber ) = TRUE then
	StdOut.WriteLine( "MINISPEC ELECTRONIC UNIT CONNECTION DONE.")
else
	StdOut.WriteLine( "---" & errMsg & "---" )
	MsgBox( errMsg )
	WScript.Quit
End If
StdOut.WriteLine ""

'###################################################################################################
'	Sets Gas Flow
'---------------------------------------------------------------------------------------------------

'Sets the given gas flow rate value to the instrument and then reads back the
'value immediately. Due to instrumental restrictions, the read back value may
'differ from the value, given in the argument.
StdOut.WriteLine ( "bvt.GasFlow: " & dFlow )
dFlow = bvt.GasFlow( dFlow )
StdOut.WriteLine( " GasFlow: " & dFlow )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "GasFlow", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""

'###################################################################################################
'	Turns Gas Flow and Heater On
'---------------------------------------------------------------------------------------------------

bOn = true
StdOut.WriteLine ( "bvt.GasFlowOn: " & bOn )
bOn = bvt.GasFlowOn( bOn)
StdOut.WriteLine( " GasFlowOn: " & bOn )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "GasFlowOn", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""


bOn = true
StdOut.WriteLine ( "bvt.HeaterOn: " & bOn )
bOn = bvt.HeaterOn( bOn )
StdOut.WriteLine( " HeaterOn: " & bOn )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "bvt.HeaterOn", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""

'###################################################################################################
'	Tune
'---------------------------------------------------------------------------------------------------

dTune = bvt.IsPIDTuneOn
StdOut.WriteLine( " IsTuneOn: " & dTune )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "IsTuneOn", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""

If	to_tune = 1 then
	dTune = bvt.PIDTuneOn( true )
	StdOut.WriteLine( " TuneOn: " & dTune )
	If	bvt.IsLastError then
		bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
		Call ErrorExit( "TuneOn", ErrMsg, ErrNo )
	End If
	if dTune = false then
		StdOut.WriteLine( " TuneOn: " & dTune )
	end if
	StdOut.WriteLine ""
	
	dTune = true
	While dTune = true
		dTune = bvt.IsPIDTuneOn
		If	bvt.IsLastError then
			bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
			Call ErrorExit( "IsTuneOn", ErrMsg, ErrNo )
		End If
		StdOut.WriteLine ""
	wend
End If

'###################################################################################################
'Gives back the actual thermo coupler temperature value immediately.
'---------------------------------------------------------------------------------------------------

StdOut.WriteLine "bvt.GetTemperature:"
dTemp = bvt.GetTemperature
StdOut.WriteLine( " GetTemperature: " & dTemp )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "bvt.GetTemperature", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""

If	starttemp = 0 then
	starttemp = dTemp
	endtemp = dTemp
End If

'###################################################################################################
'	Set temperature values
'---------------------------------------------------------------------------------------------------

'Sets the given desired temperature to the instrument and then reads back the
'value immediately. Due to instrumental restrictions, the read back value may
'differ from the value, given in the argument.
bContinue = true
if starttemp > endtemp then
	tempstep = -1*tempstep
end if
dtemp = starttemp - tempstep
while	bContinue = TRUE
	dtemp = dTemp + tempstep	'increase temperature and set
	StdOut.WriteLine ( "bvt.DesiredTemperature: " & dTemp )
	dtemp = bvt.DesiredTemperature( dTemp )
	StdOut.WriteLine( " DesiredTemperature: " & dtemp )
	If	bvt.IsLastError then
		bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
		Call ErrorExit( "bvt.DesiredTemperature", ErrMsg, ErrNo )
	End If
	StdOut.WriteLine ""

	'Gives back the desired temperature value immediately.
	StdOut.WriteLine "bvt.GetDesiredTemperature:"
	dTemp = bvt.GetDesiredTemperature
	StdOut.WriteLine( " GetDesiredTemperature: " & dTemp )
	If	bvt.IsLastError then
		bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
		Call ErrorExit( "bvt.GetDesiredTemperature", ErrMsg, ErrNo )
	End If
	StdOut.WriteLine ""
	uti.Wait waitingtime*60, "Waiting..."
	
	For Each app In appsfile
	
		if app = "" then
		
		else
	
		'###################################################################################################
		'	retrieve and build the path to the used minispec application and load application
		'---------------------------------------------------------------------------------------------------

		MinispecAppPath    = AppPath
		MinispecApplicFile = app

		errMsg = "Application Loading Failure - Program Aborted."
		StdOut.WriteLine( "LOADING MINISPEC APPLICATION..." )		' Now Load	
		if	LoadApplication( MinispecApplicFile ) = TRUE then
			StdOut.WriteLine( "MINISPEC APPLICATION LOADED." )
		else
			StdOut.WriteLine( "---" & errMsg & "---" )
			MsgBox( errMsg )
			WScript.Quit
		End If
		StdOut.WriteLine ""
		
		'###################################################################################################
		'	run the previously loaded minispec application
		'---------------------------------------------------------------------------------------------------
		errMsg = "Minispec Application Run Error - Program Aborted."
		StdOut.WriteLine( "STARTING MINISPEC APPLICATION TO RUN..." )
		if	RunApplication = TRUE then
			StdOut.WriteLine( "MINISPEC APPLICATION STARTED TO RUN." )
		else
			StdOut.WriteLine( "---" & errMsg & "---" )
			MsgBox( errMsg )
			WScript.Quit
		End If
		StdOut.WriteLine ""


		'###################################################################################################
		'	wait until measurement done
		'---------------------------------------------------------------------------------------------------
		errMsg = "Minispec Application Data Acquisition Error - Program Aborted."
		StdOut.WriteLine( "WAITING FOR DATA ACQUISITION DONE..." )
		if	WaitForDataAcqDone() = TRUE then
			StdOut.WriteLine( "MINISPEC DATA ACQUISITION DONE." )
		else
			StdOut.WriteLine( "---" & errMsg & "---" )
			MsgBox( errMsg )
			WScript.Quit
		End If
		StdOut.WriteLine ""
		
		if starttemp > endtemp then
			if	( dtemp <= endtemp ) then
					bContinue = FALSE
				End If
		else
			if	( dtemp >= endtemp ) then
				bContinue = FALSE
			End If
		end if
		
		
		end if
		
	Next
wend

'###################################################################################################
'	Turn heater off
'---------------------------------------------------------------------------------------------------

bOn = false
StdOut.WriteLine ( "bvt.HeaterOn: " & bOn )
bOn = bvt.HeaterOn( bOn)
StdOut.WriteLine( " HeaterOn: " & bOn )
If	bvt.IsLastError then
	bIsLastErr = bvt.GetLastError( ErrNo, ErrMsg )
	Call ErrorExit( "bvt.HeaterOn", ErrMsg, ErrNo )
End If
StdOut.WriteLine ""

'###################################################################################################
'	End script:
'---------------------------------------------------------------------------------------------------
StdOut.WriteLine ""
StdOut.WriteLine( "'RunApplicationsChangingTemperature.vbs' Terminated. -----------------" )



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' FUNCTIONS and SUBROUTINES (BVT):
'---------------------------------------------------------------------------------------------------

'###################################################################################################
Sub ErrorExit( strCommand, lngErrNo, strErrMsg )		'-------------------------------------------
'---------------------------------------------------------------------------------------------------
	AbortMessage = "ERROR: BVT COMMAND EXECUTION ERROR. - Script aborted."
	StdOut.WriteLine( "'" & strCommand & "' ERROR #" & lngErrNo & ": " & CHR(10) & strErrMsg )
	
	if	gblnDontAsk = FALSE then
		if	MsgBox( "'BVTFunctionCheck.vbs': Exit On Error ?", vbYesNo ) = vbNo then
			gblnExitOnError = FALSE
		else
			gblnExitOnError = TRUE
		end if
	end if
	gblnDontAsk	= TRUE	

	if	gblnExitOnError = TRUE then
		StdOut.WriteLine( "----SCRIPT ABORTED----" )
		MsgBox AbortMessage 
		WScript.Quit
	end if
end sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' FUNCTIONS and SUBROUTINES (MiniSpec):
'---------------------------------------------------------------------------------------------------

'###################################################################################################
Function ConnectMinispec( csSerno )	'---------------------------------------------------------------
'	Connects the minispec electronic unit
'	Arguments:	csSerNo		a valid serial number string to a minispec electronic unit
'							Examples: 	"ND1234", "DEMO"
'	Returns:	FALSE, if not connected, else TRUE
'---------------------------------------------------------------------------------------------------

	bConnectState = TRUE							' initializes connection state
	StdOut.WriteLine " pnmr.GetInstrumentSerialNumber:"
	csCurrSerNo = pnmr.GetInstrumentSerialNumber	' retrieves current minispec e-unit's serial no
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.GetInstrumentSerialNumber' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		bConnectState = FALSE						' not connected
	else
		StdOut.WriteLine( " Instrument SerNo	  : '" & csCurrSerNo & "'" )
		if	csCurrSerNo <> csSerno then
			bConnectState = FALSE
		End If
	End If

	if	bConnectState <> TRUE then					' re-connects in DEMO mode, if necessary
		StdOut.WriteLine ""
		StdOut.WriteLine " pnmr.ConnectInstrument:"
		bConnectState = pnmr.ConnectInstrument( csSerno )
		If	pnmr.IsLastError then
			ErrMsg = pnmr.GetLastError( ErrNo )
			StdOut.WriteLine( " 'pnmr.ConnectInstrument' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		else
			csCurrSerNo = pnmr.GetInstrumentSerialNumber
			If	pnmr.IsLastError then
				ErrMsg = pnmr.GetLastError( ErrNo )
				StdOut.WriteLine( " 'pnmr.GetInstrumentSerialNumber' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
			else
				StdOut.WriteLine( " Instrument SerNo	  : '" & csCurrSerNo & "'" )
				bConnectState = TRUE
			End If
		End If
	else
		bConnectState = TRUE
	End If

	ConnectMinispec = bConnectState
	Exit Function

End Function


'###################################################################################################
Function LoadApplication( csApplic )	'-----------------------------------------------------------
'	Loads a minispec ExpSpel application
'	Arguments:	csApplic 	a valid file name to a minispec application ('.app')
'							including its complete path 
'							Example: 	"C:\minispec Applications\fid.app"
'	Returns:	FALSE, if not loaded, else TRUE
'---------------------------------------------------------------------------------------------------
	StdOut.WriteLine " pnmr.LoadApplication:"
	StdOut.WriteLine " '" & csApplic & "'"

	pnmr.LoadApplication csApplic
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.LoadApplication' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		bLoadState = FALSE
		LoadApplication = FALSE
		Exit Function
	else
		StdOut.WriteLine " ...done."
	End If

	StdOut.WriteLine " pnmr.IsApplicationLoaded:"
	bLoaded = pnmr.IsApplicationLoaded
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.IsApplicationLoaded' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		LoadApplication = FALSE
		Exit Function
	else
		StdOut.WriteLine " ...yes."
	End If

	LoadApplication = TRUE
	Exit Function

End Function


'###################################################################################################
Function TransmitParameter( csKwd, dValue )	'---------------------------------------------------
'	Transmits an application parameter to the minispec
'	Arguments:	csKwd		a valid parameter keyword string
'							Examples: 	"SCANS", "DECDGTS"
'				dVal		a valid number, representing the value, associated with the keyword
'							Examples:	12, 32.81, -1
'	Returns:	FALSE, if not transmitted, else TRUE
'---------------------------------------------------------------------------------------------------
	'Transmission of minispec application parameters: 
	StdOut.WriteLine " pnmr.SetupApplication:"
	StdOut.WriteLine " '" & csKwd & ": " & dValue & "'"
	pnmr.SetupApplication csKwd, dValue
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.SetupApplication' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		TransmitParameter = FALSE
	else
		TransmitParameter = TRUE
	End If
	Exit Function

End Function


'###################################################################################################
Function RunApplication()	'-----------------------------------------------------------------------
'	Runs a previously loaded minispec application
'	Arguments:	none
'	Returns:	FALSE, if timed out and not started, else TRUE
'---------------------------------------------------------------------------------------------------
	StdOut.WriteLine " pnmr.IsApplicationRunning:"
	timeCnt = 0
	timeInc	= 0.5
	timeOut = 25	'sec
	while	pnmr.IsApplicationRunning = TRUE	' waits 25 sec until finnished
		StdOut.WriteLine " ...yes"
		WScript.Sleep timeInc * 1000
		timeCnt = timeCnt + timeInc
		if	( timeCnt > timeOut ) then
			StdOut.WriteLine( " 'pnmr.RunApplication' ERROR: Timed Out" )
			RunApplication = FALSE
			Exit Function
		End If
	wend
	StdOut.WriteLine " ...no"

	StdOut.WriteLine " pnmr.RunApplication:"
	pnmr.RunApplication
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.RunApplication' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		RunApplication = FALSE
		Exit Function
	else
		StdOut.WriteLine " pnmr.IsApplicationRunning:"
		timeCnt = 0
		timeInc	= 0.5
		timeOut = 10	'sec
		while	pnmr.IsApplicationRunning <> TRUE	' waits 10 sec until running
			StdOut.WriteLine " ...no"
			WScript.Sleep timeInc * 1000
			timeCnt = timeCnt + timeInc
			if	( timeCnt > timeOut ) then
				StdOut.WriteLine( " 'pnmr.RunApplication' ERROR: Timed Out" )
				RunApplication = FALSE
				Exit Function
			End If
		wend
		StdOut.WriteLine " ...yes"
	End If
	
	RunApplication = TRUE
	Exit Function

End Function


'###################################################################################################
Function WaitForDataAcqDone()	'-------------------------------------------------------------------
'	Waits until a minispec application has finished its data acquisition
'	Arguments:	none
'	Returns:	FALSE, if timed out or in case of an error, else TRUE
'---------------------------------------------------------------------------------------------------

	StdOut.WriteLine " pnmr.GetDataAcquisitionProgress:"
	' wait for data acquisition start:
	timeCnt = 0	
	timeInc	= 0.5
	timeOut = 30	'sec
	while	pnmr.GetDataAcquisitionProgress( scansToDo, scansDone ) <> TRUE
			StdOut.WriteLine( " Acq OFF " & scansDone & " of " & scansToDo & " scans done"  )
			WScript.Sleep timeInc * 1000
			timeCnt = timeCnt + timeInc
			if	( timeCnt > timeOut ) then
				StdOut.WriteLine( " 'pnmr.GetDataAcquisitionProgress' ERROR: Timed Out" )
				WaitForDataAcqDone = FALSE
				Exit Function
			End If
	wend
	totalScans = scansToDo	' scansToDo is 0 after data acquisition has been finished

	' wait for data acquisition end:
	timeCnt = 0	
	timeInc	= 1.5
	timeOut = 300	'sec
	while	pnmr.GetDataAcquisitionProgress( scansToDo, scansDone ) = TRUE
			StdOut.WriteLine( " Acq  ON " & scansDone & " of " & scansToDo & " scans done"  )
			WScript.Sleep timeInc * 1000
			timeCnt = timeCnt + timeInc
			if	( timeCnt > timeOut ) then
				StdOut.WriteLine( " 'pnmr.GetDataAcquisitionProgress' ERROR: Timed Out" )
				WaitForDataAcqDone = FALSE
				Exit Function
			End If
	wend
	StdOut.WriteLine( " Acq OFF " & totalScans & " of " & totalScans & " scans done"  )
	
	WaitForDataAcqDone = TRUE
	Exit Function

End Function


'###################################################################################################
Function WaitForDataAvail()	'-----------------------------------------------------------------------
'	Waits until a minispec application has made available its measured results
'	Arguments:	none
'	Returns:	FALSE, if timed out or in case of an error, else TRUE
'---------------------------------------------------------------------------------------------------

	StdOut.WriteLine " pnmr.IsResultAvailable:"
	timeCnt = 0	
	timeInc	= 0.5
	timeOut = 10	'sec
	while	pnmr.IsResultAvailable() <> TRUE
			StdOut.WriteLine( " ...no"  )
			WScript.Sleep timeInc * 1000
			timeCnt = timeCnt + timeInc
			if	( timeCnt > timeOut ) then
				StdOut.WriteLine( " 'pnmr.IsResultAvailable' ERROR: Timed Out" )
				WaitForDataAvail= FALSE
				Exit Function
			End If
	wend
	StdOut.WriteLine( " ...yes"  )

	WaitForDataAvail = TRUE
	Exit Function

End Function


'###################################################################################################
Function RetrieveResult( csKwd )	'---------------------------------------------------------------
'	Retrieves an application result from the minispec
'	Arguments:	csKwd		a valid result keyword string
'							Examples: 	"RVALID", "MAXIMUM"
'	Returns:	FALSE, if not received, else TRUE
'	Sets:		dRes		a globally known number, representing the value, associated with the keyword
'							Examples:	12, 32.81, -1
'	NOTE:		this is an example for an internal data exchange, without using shared files
'				see also: 'WaitForDataAvail()' and 'SignalRead()'
'---------------------------------------------------------------------------------------------------
	'minispec application results: 
	StdOut.WriteLine " pnmr.GetDataPoint: '" & csKwd & "'"

	dRes = pnmr.GetDataPoint( csKwd )
	If	pnmr.IsLastError then
		ErrMsg = pnmr.GetLastError( ErrNo )
		StdOut.WriteLine( " 'pnmr.GetDataPoint' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
		RetrieveResult= FALSE
	else
		'StdOut.WriteLine " pnmr.GetDataPoint: '" & dRes & "'"
		RetrieveResult= TRUE
	End If
	Exit Function

End Function


'###################################################################################################
Function SignalRead( csDataPairsFile )	'-----------------------------------------------------------
'	Reads the content of a data pairs file into the global buffer x, y 
'	Arguments:	csDataPairsFile 	a valid file name to a minispec data pairs file ('.dps')
'									including its complete path 
'									Example: 	"C:\minispec Applications\ND1234\fid.dps"
'	Returns:	the number of data pairs read
'				a negative value in case of an error
'	NOTE:		the standard OS file system object is used here
'	NOTE:		this is an example for a data exchange, using shared files
'				see also: 'TransmitParameter()'
'---------------------------------------------------------------------------------------------------
	Const ForReading = 1, ForWriting = 2

	StdOut.WriteLine " '" & csDataPairsFile & "'"

	Set fso = CreateObject("Scripting.FileSystemObject")			' file system object
	if	fso.FileExists( csDataPairsFile ) <> TRUE then
		StdOut.WriteLine( " 'fso.FileExists' ERROR: Does Not Exist" )
		SignalRead = -1	' file does not exitst
		Exit Function
	else
		Set fo = fso.GetFile( csDataPairsFile )						' file object
		Set ts = fso.OpenTextFile( csDataPairsFile, ForReading )	' open
	End If

	' Counts data pairs:
	dataPairsCnt = 0	' counter initialization
	bContinue	 = TRUE
	Err.Clear			' clear read errors 
	while	bContinue	' count data pairs
			On Error Resume Next
			csLine = ts.ReadLine
			if	Err.Description <> "" then		' last line
				StdOut.WriteLine( " 'fts.ReadLine' Last Line Reached" )
				ts.Close
				bContinue = FALSE
			else
				dataPairsCnt = dataPairsCnt + 1
			End If
			if	dataPairsCnt > DataPairsDim then
				DataPairsDim = DataPairsDim + DataPairsPortion
				ReDim x( DataPairsDim ), y( DataPairsDim )
				'StdOut.WriteLine( " DataPairsDim = " & DataPairsDim )
			End If
	wend

	StdOut.WriteLine( " Data Pairs Array Dim = " & DataPairsDim )
	StdOut.WriteLine( " Number Of Data Pairs = " & dataPairsCnt )
	' Read data pairs:
	Set ts = fso.OpenTextFile( csDataPairsFile, ForReading )		' re open
	Err.Clear							' clear read errors 
	cnt = 0
	while	cnt < dataPairsCnt
			csLine = ts.ReadLine
			if	Err.Description = "" then				' all right
				'StdOut.WriteLine( " ts.ReadLine: " & csLine )
				Column = Split( csLine, "	", -1, 1 )	' TAB is delimiter
				x( cnt ) = Column( 1 )
				y( cnt ) = Column( 2 )
				'StdOut.WriteLine( cnt & ": " & x(cnt) & " | " & y(cnt) )
			else										' error
				StdOut.WriteLine( " 'ts.ReadLine' ERROR: Read Error" )
				ts.Close
				SignalRead = -2
				Exit Function
			End If

			cnt = cnt + 1
	wend

	ts.Close
	SignalRead = dataPairsCnt

	Exit Function

End Function


'###################################################################################################
Function DisconnectMinispec( bClose )	'-----------------------------------------------------------
'	Releases the minispec software after script termination and closes if necessary
'---------------------------------------------------------------------------------------------------
	' wait until application has been terminated
	StdOut.WriteLine " pnmr.IsApplicationRunning:"
	timeCnt = 0	
	timeInc	= 0.5
	timeOut = 10	'sec
	while	pnmr.IsApplicationRunning() = TRUE
			StdOut.WriteLine( " ...yes"  )
			WScript.Sleep timeInc * 1000
			timeCnt = timeCnt + timeInc
			if	( timeCnt > timeOut ) then
				StdOut.WriteLine( " 'pnmr.IsApplicationRunning' ERROR: Timed Out" )
				DisconnectMinispec = FALSE
				Exit Function
			End If
	wend
	StdOut.WriteLine( " ...no"  )

	if	bClose = TRUE then
		StdOut.WriteLine " pnmr.ClosePNMR:"
		pnmr.ClosePNMR( TRUE )
		If	pnmr.IsLastError then
			ErrMsg = pnmr.GetLastError( ErrNo )
			StdOut.WriteLine( " 'pnmr.ClosePNMR' ERROR #" & ErrNo & ": " & CHR(13) & ErrMsg )
			DisconnectMinispec = FALSE
			Exit Function
		else
			StdOut.WriteLine( " PNMR Close Done." )
		End If
	End If

	DisconnectMinispec = TRUE
	Exit Function

End Function

'////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////
'// NOT CONNECTED EXAMPLES FOR OUTPUT OF DATA INTO MS-EXCEL:

'###################################################################################################
Sub ExcelOutput	'-----------------------------------------------------------------------------------
'	Output of data pairs to an Excel worksheet
'	Input:	the following the globaly declared variables:
'			DataPairsCnt (number of data pairs), x, y (abscissa and ordinate values)	
'---------------------------------------------------------------------------------------------------
	ExcelApplic.Workbooks.Add 
	ExcelApplic.Visible = True 
	for i = 1 to DataPairsCnt
		ExcelApplic.Cells(i,1).Value = x(i-1) 
		ExcelApplic.Cells(i,2).Value = y(i-1) 
	next 
End Sub


'###################################################################################################
Sub ExcelDiagram( csAbscUnit, csOrdiUnit )	'-----------------------------------------------
'	Output of a minispec signal as Excel diagram
'	Arguments:	csAbscUnit	abscissa unit string
'							Example:	"time/ms"
'				csOrdiUnit 	ordinate unit string
'							Example:	"intensity/%"
'	Input:	the following the globaly declared variables:
'			DataPairsCnt (number of data pairs), x, y (abscissa and ordinate values)
'	NOTE:	this subroutine is a slightly modified macro, automatically recorded with MS-Excel	
'
' Makro am 23.03.2000 von HMG aufgezeichnet
'---------------------------------------------------------------------------------------------------
	csRange = "A1:B" & DataPairsCnt

    ExcelApplic.Columns("A:B").Select
    ExcelApplic.Charts.Add
    ExcelApplic.ActiveChart.ChartType = 75
    ExcelApplic.ActiveChart.SetSourceData ExcelApplic.Sheets("Tabelle1").Range(csRange), 2
    ExcelApplic.ActiveChart.Location 2, "Tabelle1"
    With ExcelApplic.ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "the minispec"
        .Axes(1, 1).HasTitle = True
        .Axes(1, 1).AxisTitle.Characters.Text = csAbscUnit
        .Axes(2, 1).HasTitle = True
        .Axes(2, 1).AxisTitle.Characters.Text = csOrdiUnit
    End With
    With ExcelApplic.ActiveChart.Axes(1)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    With ExcelApplic.ActiveChart.Axes(2)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    ExcelApplic.ActiveChart.HasLegend = False
    ExcelApplic.ActiveChart.PlotArea.Select
    With ExcelApplic.Selection.Border
        .ColorIndex = 16
        .Weight = 2
        .LineStyle = 1
    End With
    ExcelApplic.Selection.Interior.ColorIndex = -4142
    ExcelApplic.ActiveChart.SeriesCollection(1).Select
    With ExcelApplic.Selection.Border
        .ColorIndex = 57
        .Weight = -4138
        .LineStyle = 1
    End With
    With ExcelApplic.Selection
        .MarkerBackgroundColorIndex = -4142
        .MarkerForegroundColorIndex = -4142
        .MarkerStyle = -4142
        .Smooth = False
        .MarkerSize = 3
        .Shadow = False
    End With
    ExcelApplic.ActiveChart.ChartArea.Select
End Sub
