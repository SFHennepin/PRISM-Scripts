Option Explicit
Dim beta_agency
'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF 'copy/paste - this is the custom function script that needs to be palced on top of your dialog editor' 
'DIM functions for dialog
Dim emancipation_Review_dialog, prism_case_number, child_name, expected_date_of_graduation, worker_signature, buttonpressed 'this declares your dialog'
'DIALOGS emancaiption_review 
BeginDialog emancipation_review_dialog, 0, 0, 251, 125, "Dialog"
  EditBox 75, 20, 50, 15, prism_case_number
  EditBox 55, 40, 140, 15, child_name
  EditBox 105, 60, 60, 15, expected_date_of_graduation
  EditBox 95, 80, 95, 15, worker_signature
ButtonGroup ButtonPressed
    OkButton 125, 105, 50, 15
    CancelButton 190, 105, 50, 15
  Text 5, 5, 75, 10, "Emancipation Review"
  Text 5, 25, 70, 10, "Prism Case Number"
  Text 5, 45, 50, 10, "Child Name"
  Text 5, 65, 100, 10, "Expected Date of Graduation:"
  Text 5, 85, 40, 10, "worker_signature"
EndDialog

EMConnect "" 'connects to bluescript' 

DO
Dialog emancipation_review_dialog
IF worker_signature = "" THEN MsgBox "You must sign!"
LOOP UNTIL worker_signature <> ""

Dialog emancipation_review_dialog

CALL navigate_to_PRISM_screen("CAAD")
PF5
EMWriteScreen "A", 3, 29
transmit 

'Writing the case note
EMWriteScreen "FREE", 4, 54
EMSetCursor 16, 4
CALL write_variable_in_caad("Emancipation Review")
CALL write_bullet_and_variable_in_caad("Child Name", Child_Name)
CALL write_bullet_and_variable_in_caad("Expected Date of Graduation", Expected_Date_of_Graduation)
CALL write_variable_in_caad(worker_signature) 
 

