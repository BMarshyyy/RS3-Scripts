'==============================================================================================================
'TEST MODE - REROUTES SAVED FILES
TESTING_MODE = TRUE

'==============================================================================================================
name_of_script = "Change - Address Change.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 10
STATS_denomination = "CUSTOM"
'END OF STATS BLOCK===========================================================================================

'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING FUNCTIONS LIBRARY FROM REPOSITORY===========================================================================
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
'END FUNCTIONS LIBRARY BLOCK=============================================================================================


' ENSURE WORKERS ARE CONNECTED TO THE NETWORK DRIVE
failed_msg = "You do not currently have a connection to the financial drive." & vbcr & vbcr & "- Try the following to resolve the issue - " & _
              vbcr & "1.) VPN users - Log out of windows and ensure you connect to the VPN before logging back in." &_
              vbcr & "2.) Ensure you have a connection to the financial folder." &_
              vbcr & "3.) If none of these resolve your issue, contact the Benton script admin." &_
			  vbcr & "The script will not work until the issue is resolved."

CALL FolderExists("", CStr(false_msg))    

'CONNECTING TO MAXIS
EMConnect("")


' VERIFY FILE SYSTEM CONNECTIOn
Call FolderExists("", "Connection to the Benton file system failed. Please contact your script administrator.")

' VERIFYING WORKER IS IN MAXIS...
Call check_for_MAXIS(false)

' ENSURE SCRIPT IS RUN FROM SELF MENU
Call back_to_SELF()

' LOCATES CASE NUMBER / BENEFIT MONTH
Call MAXIS_case_number_finder(MAXIS_case_entry)
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" Then
	CALL find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" Then CALL find_variable("Benefit Period (MM YY): " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
Else
	CALL find_variable("Month: ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" Then CALL find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
END IF


' MAXIS OR METS SELECTION DIALOG / CALL DIALOG1
BeginDialog Dialog1, 0, 0, 191, 115, "Dialog"
  Text 10, 10, 80, 15, "MAXIS Case Number:"
  Text 10, 40, 70, 15, "METS Case Number:"
  Text 10, 70, 40, 10, "MNsure ID:"
  EditBox 95, 5, 55, 15, MAXIS_case_entry
  EditBox 95, 35, 55, 15, METS_case_entry
  EditBox 95, 65, 55, 15, MNsure_ID_entry
  ButtonGroup ButtonPressed
    PushButton 20, 95, 50, 15, "OK", ok_button
    PushButton 125, 95, 45, 15, "Cancel", cancel_button
EndDialog


' FIRST CASE NUMBER DIALOG W/ ERROR HANDLING
DO
	err_msg = ""

	Dialog(Dialog1)

    If ButtonPressed = cancel_button Then script_end_procedure("~PT: User Pressed cancel")
    If MAXIS_case_entry = "" AND METS_case_entry = "" AND MNsure_ID_entry = "" Then err_msg = err_msg & vbCr & "Please enter a valid number"
    If (MAXIS_case_entry <> "" AND len(MAXIS_case_entry) > 8) OR (METS_case_entry <> "" AND len(METS_case_entry) > 10) OR (MNsure_ID_entry <> "" AND len(MNsure_ID_entry) > 8) Then err_msg = err_msg & vbCr & "Please enter a valid number."
    If (MAXIS_case_entry <> "" AND IsNumeric(MAXIS_case_entry) = False) OR (METS_case_entry <> "" AND IsNumeric(METS_case_entry) = False) OR (MNsure_ID_entry <> "" AND IsNumeric(MNsure_ID_entry) = False) Then err_msg = err_msg & vbCr & "Please enter a valid number."
    
    If err_msg <> "" Then MsgBox(err_msg)
LOOP UNTIL err_msg = ""

' CREATED HH MB ARRAY - ALSO FOR EMPTY 
hh_member_array = Array()

' IF MAXIS - NAVIGATE/COLLECT CLIENT DATA 
If  ButtonPressed <> cancel_button AND MAXIS_case_entry <> "" Then
    call navigate_to_MAXIS_screen("SELF", "")
    EMWriteScreen "stat", 16, 043
    EMWriteScreen MAXIS_case_entry, 18, 043 
    EMWriteScreen "memb", 21, 070
    transmit

    EMReadScreen err_check, 4, 02, 052
    If err_check = "ERRR" Then Transmit

    ' HANDLING PRIV CASES
    EMReadScreen priv_check, 10, 24, 014
    IF priv_check <> "PRIVILEGED" THEN

        ' HANDING STAT MSG UPDATES
        EMReadScreen stat_err1, 05, 08, 27
        IF stat_err1 = "Note:" Then transmit

        fname_array = Array()
        lname_array = Array()
        birth_date_array = Array()
        social_num_array = Array() 
        hh_members_text = "" 
        
        DO  
            ' COLLECT CLIENT INFORMATION INTO ARRAYS (HH NUMBER, FNAME, LNAME, SSN, DOB)
            ' COMP PAGE TO NAVIGATE
            EMReadScreen memb_page1, 2, 02, 072
            EMReadScreen memb_page_end, 2, 02, 078

            'COLLECTING FNAM LNAME AND REF NUMBER
            EMReadScreen ref_num, 2, 04, 033
            EMReadScreen f_name, 12, 06, 063
            EMReadScreen l_name, 25, 06, 030
            EmReadScreen birth_date, 10, 08, 42
            dob = replace(birth_date, " ", "/") 
            EmReadScreen social_num, 11, 07, 42
            ssn = replace(social_num, " ", "-")

            'PARSING MEMBER NAME/NUMBER/CASE NAME
            member = ref_num & " " & replace(f_name, "_", "") & " " & replace(l_name, "_", "")
            hh_members_text = hh_members_text & member & chr(9)
            If memb_page1 = " 1" Then case_name = l_name

            ReDim Preserve hh_member_array(UBound(hh_member_array) + 1)
            hh_member_array(UBound(hh_member_array)) = member

            ReDim Preserve fname_array(UBound(fname_array) + 1)
            fname_array(UBound(fname_array)) = f_name

            ReDim Preserve lname_array(UBound(lname_array) + 1)
            lname_array(UBound(lname_array)) = l_name

            ReDim Preserve birth_date_array(UBound(birth_date_array) + 1)
            birth_date_array(UBound(birth_date_array)) = dob

            ReDim Preserve social_num_array(UBound(social_num_array) + 1)
            social_num_array(UBound(social_num_array)) = ssn

            IF Trim(memb_page1) <> Trim(memb_page_end) Then Transmit
        LOOP UNTIL Trim(memb_page1) = Trim(memb_page_end)

        ' REMOVING LAST RETURN FROM MSG TO AVOID EXTRA COMBOBOX ENTRY 
        hh_members_text = Left(hh_members_text, Len(hh_members_text) -1)

    END IF

End if



' SECOND CASE NUMBER DIALOG W/ ERROR HANDLING
exit_dialog = False
DO
	err_msg = ""

    BeginDialog Dialog1, 0, 0, 520, 440, "Address or Phone Change"
        ButtonGroup ButtonPressed
            PushButton 185, 420, 50, 15, "Cancel", cancel_button
            PushButton 295, 420, 50, 15, "Next", next_button
            If UBound(hh_member_array) <> -1 Then PushButton 295, 10, 60, 15, "Update", update_hh_memb 
        If UBound(hh_member_array) <> -1 Then ComboBox 165, 10, 95, 20, hh_members_text, member_combo 
        Text 15, 40, 50, 10, "Call Taken By:"
        Text 20, 80, 40, 10, "First Name:"
        Text 20, 100, 45, 10, "Last Name:"
        If MAXIS_case_entry <> "" THEN
            Text 250, 40, 70, 15, "METS Case Number:"
            Text 280, 70, 40, 15, "MNsure ID:"
        END IF
        Text 235, 100, 85, 10, "Which change occured?"
        Text 20, 60, 45, 10, "Case Name:"
        Text 50, 150, 110, 10, "Date Moved (Month/Date):"
        Text 225, 150, 120, 10, "Who Moved:"
        Text 360, 150, 130, 10, "Is this move permanent or temporary?"
        Text 50, 190, 75, 10, "Residential Address:"
        Text 225, 190, 30, 10, "County:"
        Text 50, 230, 55, 10, "Living Situation:"
        Text 225, 225, 120, 15, "Is there a change in housing or utility costs?"
        Text 360, 225, 70, 10, "If so, what changes?"
        Text 360, 280, 95, 15, "Should any phone number be removed on the case?"
        Text 50, 285, 30, 10, "Phone 1: "
        Text 50, 325, 30, 10, "Phone 2:"
        Text 225, 280, 45, 10, "Phone Type:"
        Text 225, 325, 50, 10, "Phone Type:"
        Text 360, 325, 95, 10, "Preferred method of contact:"
        Text 50, 365, 125, 20, "If last residence was a facility, what was the discharge date?"
        Text 225, 370, 75, 10, "Additional Comments:"
        EditBox 70, 40, 85, 15, call_taken_by
        EditBox 70, 60, 85, 15, case_name_entry
        EditBox 70, 80, 85, 15, first_name
        EditBox 70, 100, 85, 15, last_name
        If MAXIS_case_entry <> "" THEN
            EditBox 325, 40, 60, 15, METS_case_entry
            EditBox 325, 70, 60, 15, MNsure_ID_entry
        END IF
        DropListBox 325, 100, 100, 15, chr(9) + "No move, Phone Only"+chr(9)+"Mailing Only"+chr(9)+"Mailing and Residential"+chr(9)+"Out of State", move_list
        EditBox 50, 165, 100, 15, date_moved_entry
        EditBox 225, 165, 100, 15, who_moved_entry
        DropListBox 360, 165, 70, 20, chr(9) + "permanent"+chr(9)+"Temporary", move_duration_list
        EditBox 50, 205, 100, 15, res_address_entry
        EditBox 225, 205, 100, 15, county_entry
        DropListBox 50, 245, 155, 15, chr(9) + "01 - Own housing/Lease/Mortgage or roommate"+chr(9)+"02 - Family/Friend due to econmic hardship"+chr(9)+"03 - Service provider- Foster care group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing home"+chr(9)+"05 - Jail/Prison/Juvenile detention center"+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown", living_situation_list
        DropListBox 225, 245, 70, 15, chr(9) + "Yes"+chr(9)+"No", utility_change_list
        EditBox 360, 240, 100, 15, costs_change_entry
        EditBox 50, 300, 80, 15, phone_one_entry
        DropListBox 225, 300, 70, 15, chr(9) + "Cell"+chr(9)+"Home"+chr(9)+"Other", phone_one_list
        EditBox 360, 300, 100, 15, phone_remove_entry
        EditBox 50, 340, 80, 15, phone_two_entry
        DropListBox 225, 340, 70, 15, chr(9) + "Cell"+chr(9)+"Home"+chr(9)+"Other", phone_two_list
        EditBox 360, 340, 105, 15, contact_method_entry
        EditBox 50, 385, 120, 15, discharge_date_entry
        EditBox 225, 385, 215, 15, add_comments_entry
        GroupBox 15, 135, 490, 130, "Address Change"
        GroupBox 15, 270, 490, 140, "Phone Change"

    EndDialog

	Dialog(Dialog1)

    IF ButtonPressed = cancel_button Then
        script_end_procedure("~PT: User Pressed cancel")
    ElseIf ButtonPressed = update_hh_memb then
        For i = 0 To UBound(hh_member_array) 

            temp_first_name = replace(fname_array(i), "_", "")
            split_combo = split(member_combo, " ", 3)

            if temp_first_name = split_combo(1) Then
                index_select = i
                Exit For 
            End If
        Next

        first_name = replace(fname_array(index_select), "_", "")
        last_name = replace(lname_array(index_select), "_", "")
        dob_entry = birth_date_array(index_select)
        ssn_entry = social_num_array(index_select)   

        case_name_entry = last_name

    ELSEIf ButtonPressed = next_button then  
        IF MAXIS_case_entry = "" AND METS_case_entry = "" AND MNsure_ID_entry = "" Then err_msg = err_msg & vbCr & "Please enter a valid number"
        IF call_taken_by = "" Then err_msg = err_msg & vbCR & "Check to ensure Call Taken By is completed."
        IF case_name_entry = "" Then err_msg = err_msg & vbCR & "Check to ensure Case Name is completed."
        IF first_name = "" OR last_name = "" Then err_msg = err_msg & vbCR & "Check to ensure first and last name are completed."
        IF move_list = "" Then
            err_msg = err_msg & vbCR & "Select which change occured."
        elseif move_list <> "No move, Phone Only" Then 
            IF date_moved_entry = "" OR who_moved_entry = "" OR res_address_entry = "" OR county_entry = "" Then err_msg = err_msg & vbCR & "Check to ensure all address change entries are completed."
            IF move_duration_list = "" Then err_msg = err_msg & vbCR & "Select if the move is temporary or permanent."
            IF living_situation_list = "" Then err_msg = err_msg & vbCR & "Select a living situation."
        END If

        IF utility_change_list = "Yes" Then 
            IF costs_change_entry = "" Then err_msg = err_msg & vbCR & "Explain what changes have been made for housing or utiltiy costs."
        END IF

        If move_list = "No move, Phone Only" Then
            IF len(phone_one_entry) <> 12 Then err_msg = err_msg & vbCR & "Please check the format for Phone 1 is as follows: 555-555-5555."
            IF len(phone_two_entry) <> 0 Then
                IF len(phone_two_entry) <> 12 Then err_msg = err_msg & vbCR & "Please check the format for Phone 2 is as follows: 555-555-5555."
            END IF
            IF contact_method_entry = "" Then err_msg = err_msg & vbCR & "Please state a preferred method of contact."
        End if

        IF err_msg = "" Then exit_dialog = True
    END IF

    IF err_msg <> "" Then MsgBox(err_msg)

LOOP UNTIL err_msg = "" and exit_dialog = True


' FILL IN WORD DOCUMENT TEMPLATE WITH DATA 
set objword = CreateObject("Word.Application")
objword.Visible = True

Set objDoc = objWord.Documents.Open("", , True)
With objDoc
    .FormFields("call_taken_by").Result = call_taken_by
    .FormFields("maxis_case_number").Result = MAXIS_case_number
    .FormFields("mets_case_number").Result = METS_case_entry
    .FormFields("mnsure_id").Result = MNsure_ID_entry
    .FormFields("case_name").Result = case_name_entry
    .FormFields("which_move").Result = move_list
    .FormFields("date_moved").Result = date_moved_entry
    .FormFields("Who_moved").Result = who_moved_entry
    .FormFields("residential_address").Result = res_address_entry
    .FormFields("addr_county").Result = county_entry
    .FormFields("temp_or_perm").Result = move_duration_list
    .FormFields("living_situation").Result = living_situation_list
    .FormFields("change_util_cost").Result = utility_change_list
    .FormFields("change_housing_cost").Result = costs_change_entry
    .FormFields("phone_one").Result = phone_one_entry
    .FormFields("phone_one_type").Result = phone_one_list
    .FormFields("phone_two").Result = phone_two_entry
    .FormFields("phone_two_type").Result = phone_two_list
    .FormFields("remove_number").Result = phone_remove_entry
    .FormFields("pref_contact_method").Result = contact_method_entry
    .FormFields("date_discharge").Result = discharge_date_entry
    .FormFields("additional_comments").Result = add_comments_entry

    objDoc.Application.DisplayAlerts = False 'Prevents prompts from appearing

END WITH

script_end_procedure("Success! Your script is complete.")




