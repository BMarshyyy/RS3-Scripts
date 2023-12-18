'====================================================================================================================
'                                               BENTON INTERVIEW TEXT REMINDER
'                                          Developed by Brenton Marshik(Benton County)
'                                                    Last Updated 09/2023
'====================================================================================================================

'Required for statistical purposes==========================================================================================
name_of_script = "Interview-Reminder"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "TEXTING"        'C is for each case
'END OF stats block=========================================================================================================


' DIALOG FOR INFO ENTRY
' ===============================================================================================================================================
' ===============================================================================================================================================


BeginDialog Dialog1, 0, 0, 211, 225, "Text Message Information"
  ButtonGroup ButtonPressed
    OkButton 35, 195, 50, 15
    CancelButton 135, 195, 50, 15
  Text 20, 40, 75, 10, "Client Phone Number :"
  Text 35, 70, 60, 10, "Date of Interview :"
  Text 35, 100, 60, 10, "Time of Interview :"
  Text 10, 130, 85, 10, "Worker Callback Number :"
  Text 100, 40, 5, 10, "("
  Text 125, 40, 5, 10, ")"
  Text 155, 40, 5, 10, "-"
  Text 100, 130, 5, 10, "("
  Text 125, 130, 5, 10, ")"
  Text 155, 130, 5, 10, "-"
  Text 35, 160, 60, 10, "Worker Signature:"
  Text 45, 10, 50, 10, "Case Number :"
  EditBox 100, 10, 60, 15, MAXIS_case_number
  EditBox 103, 40, 20, 15, client_areacode
  EditBox 130, 40, 20, 15, client_prefix
  EditBox 165, 40, 25, 15, client_line_num
  EditBox 100, 70, 55, 15, inter_date
  EditBox 100, 100, 40, 15, inter_time
  DropListBox 150, 100, 30, 15, "-"+chr(9)+"AM"+chr(9)+"PM", time_dropdown
  EditBox 103, 130, 20, 15, worker_areacode
  EditBox 130, 130, 20, 15, worker_prefix
  EditBox 165, 130, 25, 15, worker_line_num
  EditBox 100, 160, 45, 15, worker_signature
EndDialog

BeginDialog Verification_Dialog, 0, 0, 281, 115, "!! MAXIS Coding Issue !!"
  ButtonGroup ButtonPressed
    OkButton 120, 85, 50, 15
  Text 15, 15, 255, 10, "Phone numbers on the ADDR panel are not coded to allow for text messaging."
  Text 15, 30, 260, 10, "Confirm with the application that the client can be texted and code it in MAXIS."
  Text 95, 65, 110, 10, "Press ""OK"" when completed."
EndDialog

'''''''' - ALL INFORMATION MUST BE VALIDATED - ''''''''''''''''''''

' connects to BlueZone and brings it forward
EMConnect ""
EMFocus

' Checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(false)

' Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)

' Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
ELSE
	CALL find_variable("Month: ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
END IF

' CALL DIALOG AND VALIDATE ENTRIES IN DIALOG BOX
DO
	err_msg = ""

	Dialog(Dialog1)  ' put dialog run here
  IF len(client_areacode) <> 3 or IsNumeric(client_areacode) = False THEN err_msg = err_msg & vbCr & "Verify client's three digit area code." 
  IF len(client_prefix) <> 3 or IsNumeric(client_prefix) = False THEN err_msg = err_msg & vbCr & "Verify client's middle three digits of phone number." 
  IF len(client_line_num) <> 4 or IsNumeric(client_line_num) = False THEN err_msg = err_msg & vbCr & "Verify client's last four digits of phone number." 
  IF IsDate(inter_date) = False or InStr(inter_date, "/") = 0 THEN err_msg = err_msg & vbCr & "Verify date is correct."
  If InStr(inter_time, ":") <> 2 THEN
    If InStr(inter_time, ":") <> 3  THEN err_msg = err_msg & vbCr & "Verify interview time in hh:mm format." 
  End IF
  If len(inter_time) < 4 or len(inter_time) > 5 THEN err_msg = err_msg & vbCr & "Verify interview time is correct." 
  IF time_dropdown <> "AM" and time_dropdown <> "PM" THEN err_msg = err_msg & vbCr & "You must select either AM or PM."
  IF len(worker_areacode) <> 3 or IsNumeric(worker_areacode) = False THEN err_msg = err_msg & vbCr & "Verify worker's three digit area code."
  IF len(worker_prefix) <> 3 or IsNumeric(worker_prefix) = False THEN err_msg = err_msg & vbCr & "Verify worker's middle three digits of phone number."
	IF len(worker_line_num) <> 4 or IsNumeric(worker_line_num) = False THEN err_msg = err_msg & vbCr & "Verify worker's last four digits of phone number."

  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""


' NAVIGATE TO ADDR PANEL TO CHECK IF TEXTS CAN BE SENT 
call navigate_to_MAXIS_screen("STAT", "ADDR")

show_error = True
Do 
  For i = 16 to 18
    EMReadScreen can_text_client, 1, i, 066
    If can_text_client = "Y" THEN show_error = False
  Next
  If show_error = True Then Dialog(Verification_Dialog)
  If ButtonPressed <> -1 THEN script_end_procedure("~PT: user pressed cancel")
Loop Until show_error = False 


' CREATE TEXT FILE
' ===============================================================================================================================================
' ===============================================================================================================================================

client_full_number = client_areacode & client_prefix & client_line_num
worker_full_number = worker_areacode & worker_prefix & worker_line_num
worker_number_dashes = worker_areacode & "-" &  worker_prefix & "-" & worker_line_num
file_name = worker_full_number & "_" & replace(FormatDateTime(Now, 2), "/", "") & "_" & replace(FormatDateTime(Now, 4), ":", "")


Set objFSO=CreateObject("Scripting.FileSystemObject")
file_path = ""
Set objFile = objFSO.CreateTextFile(file_path, True)
  with objFile
    .WriteLine(client_full_number & "," & inter_date & "," & inter_time & " " & time_dropdown & "," & worker_number_dashes) ' Phone number to start - lindsey has to figure out what can go in here
    .Close()
  END WITH

' CASE NOTE 
' ===============================================================================================================================================
' ===============================================================================================================================================

call start_a_blank_case_note
call write_variable_in_CASE_NOTE("*** Interview Text Reminder Sent ***")
call write_variable_in_CASE_NOTE("  ")
call write_variable_in_CASE_NOTE("*Sent: " & FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 4))
call write_variable_in_CASE_NOTE("*Text Content Sent - - - ")
call write_variable_in_CASE_NOTE("""Benton County: You have been scheduled for a phone interview on " & inter_date & " at " & inter_time & " " & time_dropdown & "." &_ 
                                 " Please watch for a call from " & worker_number_dashes & " at your scheduled time. DO NOT REPLY to this text - this inbox is not monitored.""")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

' ===============================================================================================================================================
' ===============================================================================================================================================


script_end_procedure("Success! The apt reminder has been sent to " & client_areacode & client_prefix & client_line_num)