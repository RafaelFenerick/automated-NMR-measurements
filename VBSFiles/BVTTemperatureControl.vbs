'---------------------------------------------------------------------------------------------------
'	'BVTTemperatureControl.vbs'
' 	
'	Changes BVT temperature either reading from file or with specified start, end and step values.
'
'---------------------------------------------------------------------------------------------------


'###################################################################################################
' Getting temperature values
'---------------------------------------------------------------------------------------------------
'Specified values
starttemp = 300
tempstep = 5
endtemp = 310
waitingtime = 5
dFlow = 2000

'Reading input file (Comment to use specified values)
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile ("input.txt", 1)
row = 0
Do Until file.AtEndOfStream
	if row = 0 then
		starttemp = CDbl(file.Readline)
	end if
	if row = 1 then
		tempstep = CDbl(file.Readline)
	end if
	if row = 2 then
		endtemp = CDbl(file.Readline)
	end if
	
	row = row + 1
Loop

'###################################################################################################
'The 'WinAcquisit.Embedding' Server of the WinAcquisit Software offers the methods for configuring 
'its wake up behaviour.
'---------------------------------------------------------------------------------------------------
Dim StdIn, StdOut
Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut
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
'	Header Output
'---------------------------------------------------------------------------------------------------
StdOut.WriteLine ""
StdOut.WriteLine "-----------------------------------------------------"
StdOut.WriteLine "            'BVTTemperatureControl.vbs'			   "
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
'	Set temperature values
'---------------------------------------------------------------------------------------------------

'Sets the given desired temperature to the instrument and then reads back the
'value immediately. Due to instrumental restrictions, the read back value may
'differ from the value, given in the argument.
bContinue = true
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
	if	( dtemp >= endtemp ) then
			bContinue = FALSE
		End If
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
StdOut.WriteLine( "'BVTTemperatureControl.vbs' Terminated. -----------------" )




'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' FUNCTIONS and SUBROUTINES:
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


