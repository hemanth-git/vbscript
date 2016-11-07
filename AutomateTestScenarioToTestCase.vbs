 'Precondition before processing the documents
 '--------------------------------------------------------
 'preconditions for Test Scenario document
 'this document works fine with the NTIS_ALL and developed according to it and may be same for all with 
 'provider in column-> K
 'DOS in column-> L
 'modifier in column-> M
 'POS in column-> O
 'assumed that the DOB and GENDER is any and no reference given
			'dob = "any"
			'gender ="any"
 'Expected Result in deny, paid or paid, deny formate
 'this works well for paid claims and shouldn't process and for claim should deny.
 '
 '
 '
 
 Set objXLScenario = CreateObject("Excel.Application")
 Set objWBScenario = objXLScenario.WorkBooks.Open("C:\Users\kakarlah\Desktop\Automate Data\TestScenarios.xlsx")
 
 
 Set objXLTCases = CreateObject("Excel.Application")
 Set objWBTCases = objXLTCases.WorkBooks.Open("C:\Users\kakarlah\Desktop\Automate Data\TestCases.xlsx")

 'finding number of scenarios
 '----------------------------
	'nooftestScenarios = objXLScenario.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	Const xlUp = -4162
	'This is the Rule or sheet name to be given to process.
	ruleName="NTIS_ALL" 
	set objWSScenario = objWBScenario.Sheets.Item(ruleName)
	Set ws = objWBScenario.Worksheets(ruleName)
	With ws
		nooftestScenarios = .Range("A" & .Rows.Count).End(xlUp).Row
	End With
	msgbox "Total number of scenarios"&nooftestScenarios
	'heading
	'--------------------
		objXLTCases.Cells(1,1).value = "Test case name"
		objXLTCases.Cells(1,2).value = "Test Description"
		objXLTCases.Cells(1,3).value = "Steps"
		objXLTCases.Cells(1,4).value = "steps description"
	
	j=1
	for i = 2 to nooftestScenarios
		'test case name to generate automatically
		j=j+1
		objXLTCases.Cells(j,1).value = "CXT_PKG2_"&ruleName&"_TC00"&(i-1)
		objXLTCases.Cells(j,2).value = objWSScenario.Cells(i,2).value
		'step 1
		'---------------
			objXLTCases.Cells(j,3).value ="step 1"
			objXLTCases.Cells(j,4).value ="Login to Facets Application and From the Facets Application List, Select Subscriber/Family"
			objXLTCases.Cells(j,5).value ="Subscriber/Family tab opens."
		'step 2
		'---------------
			lob = objWSScenario.Cells(i,10).value
			'assumed that the DOB and GENDER is any and no reference given
			dob = "any"
			gender ="any"
			j=j+1
			objXLTCases.Cells(j,3).value ="step 2"
			objXLTCases.Cells(j,4).value = "Create a subscriber with the following parameters"&vbcrlf &" LOB= "&lob &vbcrlf&" DOB= "&dob &vbcrlf&" DOB= "&gender
			objXLTCases.Cells(j,5).value = "Subscriber is successfully created with the predefined parameters."
		'step 3
		'----------------
			claimtype = objWSScenario.Cells(i,9).value
			j=j+1
			objXLTCases.Cells(j,3).value ="step 3"
			if claimtype = "Medical" then
				objXLTCases.Cells(j,4).value ="From the Facets Application List, Select Claims Processing + ITS \ Med. Claims Processing + ITS"
				objXLTCases.Cells(j,5).value ="The Med. Claims Processing + ITS Application opens"
			else 
				objXLTCases.Cells(j,4).value ="From the Facets Application List, Select Claims Processing + ITS \ Hos. Claims Processing + ITS"
				objXLTCases.Cells(j,5).value ="The Hos. Claims Processing + ITS Application opens"
			end if
			
		'step 4
		'-----------------
			provider = objWSScenario.Cells(i,11).value
			j=j+1
			objXLTCases.Cells(j,3).value ="step 4"
			objXLTCases.Cells(j,4).value ="On the Indicative Screen, Make the following Entries:Subscriber Id from Step 2"&vbcrlf&" Provider= "&provider
			objXLTCases.Cells(j,5).value ="User should be able to provide the details successfully"
			
		'step 5
		'------------------
		'same for all
			j=j+1
			objXLTCases.Cells(j,3).value = "step 5"
			objXLTCases.Cells(j,4).value = "Select the Line Items sheet"
			if claimtype = "Medical" then
				objXLTCases.Cells(j,5).value = "The Line Items sheet opens in the Med. Claims Processing + ITS application"
			else 
				objXLTCases.Cells(j,5).value = "The Line Items sheet opens in the Hos. Claims Processing + ITS application"
			end if
		'step 6
		'-------------------
			dos = objWSScenario.Cells(i,12).value
			pos = objWSScenario.Cells(i,15).value
			diagnosis = "any"
			procedurecodes = objWSScenario.Cells(i,3).value
			j=j+1
			objXLTCases.Cells(j,3).value = "step 6"
			objXLTCases.Cells(j,4).value = "On the Line Items sheet, make the following entries:"&vbcrlf&"DOS : "&dos&vbcrlf&"POS : "&pos&vbcrlf&procedurecodes&vbcrlf&"NOTE: Repeat as needed for multiple line items on a single claim. Total of Line Item Charges must match Total Charge entry"
			objXLTCases.Cells(j,5).value = "User should be able to provide the details successfully"
		'step 7
		'--------------------
		'same for all
			j=j+1
			objXLTCases.Cells(j,3).value = "step 7"
			objXLTCases.Cells(j,4).value = "From the Menu Bar, select File \ Process (or F3) for adjudication of claim"
			objXLTCases.Cells(j,5).value = "Claim adjudicates"
			
		'step 8
		'---------------------
			j=j+1
			objXLTCases.Cells(j,3).value = "step 8"
			excdcode = objWSScenario.Cells(i,5).value
			validate = objWSScenario.Cells(i,4).value
			if excdcode = "" then
				if validate = "Claims shouldn't process" then
					objXLTCases.Cells(j,4).value = "Claims shouldn't process"
					objXLTCases.Cells(j,5).value = "Claims shouldn't process"
				else
					objXLTCases.Cells(j,4).value = "Validate that the claim line is Paid"
					objXLTCases.Cells(j,5).value = "Claim should be Paid i.e., the total charged amount is not equal to disallowed amount"

				end if
			else	
				excdmessage = objWSScenario.Cells(i,6).value
				clinicaledit = objWSScenario.Cells(i,7).value
				warningmessage = objWSScenario.Cells(i,8).value
				'assumed that the paid is mentioned first and deny as second
				validateSplit = split(validate,"paid")
				paidProcedurecodes=""
				if ubound(validateSplit)>1 then
					for each sentenceValidate in validateSplit
						if InStrRev(sentenceValidate,"denied") or InStrRev(sentenceValidate,"deny")  then
							paidProcedurecodes = paidProcedurecodes+"The line with procedure code " &sentenceValidate &" i.e.,Total charged amount =  Total disallowed amount"&vbcrlf
						else
							paidProcedurecodes = paidProcedurecodes+"The line with procedure code " &sentenceValidate&"is paid ie. total charged amount not equal to disallowed amount"&vbcrlf
						end if
					next
				else
					'assumed that the denied is mentioned first and paid as second
					validateSplit = split(validate,"deny")
					for each sentenceValidate in validateSplit
						if InStrRev(sentenceValidate,"paid") or InStrRev(sentenceValidate,"pay")  then
							paidProcedurecodes = paidProcedurecodes+"The line with procedure code " &sentenceValidate&" ie. total charged amount not equal to disallowed amount"&vbcrlf
						else
							paidProcedurecodes = paidProcedurecodes+"The line with procedure code " &sentenceValidate &"is denied i.e.,Total charged amount =  Total disallowed amount"&vbcrlf
						end if
					next
				end if
				
				objXLTCases.Cells(j,4).value = "Validate the line with procedure "&validate&vbcrlf&"EXCD Codes: "&excdcode&vbcrlf&"EXCD message: "&excdmessage&vbcrlf&"Clinical Edits: "&clinicaledit&vbcrlf&"Warning message: "&warningmessage
				objXLTCases.Cells(j,5).value = paidProcedurecodes
			
			end if
			
		'step 9
		'--------------------
		'same for all
			j=j+1
			objXLTCases.Cells(j,3).value = "step 9"
			objXLTCases.Cells(j,4).value = "Save the claim and note the ID for reference"
			objXLTCases.Cells(j,5).value = "Claim should be saved & Claim ID should be noted successfully."
		
		
	next
'step 2
 
 msgbox("completed success")
 objWBScenario.Save
 objWBScenario.Close
 objXLScenario.Quit
 
 objWBTCases.Save
 objWBTCases.Close
 objXLTCases.Quit