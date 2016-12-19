<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' VIATEQ CORPORATION
' FORM DEVELOPMENT FILE
' Receives form data and processes it for distribution.
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim error, justadd, justread, justwrite, newline, ok
Dim knt, f, emsg, mail_to, r, o, c, other, html, html_end, txt, etxt
Dim greetings
Dim htmlheader
Dim htmltail
Dim htmlsubject
Dim htmltitle
Dim part1, part2, part3, part4, part5, part6, part7, part8, part9
Dim member_info
Dim txt_F_Allegations1
Dim region
Dim region_title
Dim region_email
Dim F_Race

'----------------------------------------
'  Initialize Parts/Sections
'----------------------------------------

region 		= 0
justadd 	= 8
justread 	= 1
justwrite 	= 2
error 		= 0
knt 		= 0
ok 		= 0
mail_to 	= "ddunston@viateq.com"

part1 		= ""
part2 		= ""
part3 		= ""
part4 		= ""
part5 		= ""
part6 		= ""
part7 		= ""
part8 		= ""
part9 		= ""
greetings 	= ""
newline 	= vbCrLF
hline 		= vbNewLine  
emsg 		= ""
etxt		= ""

txt_F_Allegations1 = ""

F_Race 		= ""


'For Each f In Request.Form
'	If Request.Form(f) = "" Then 
'		error = 0
'	End If
'Next

'Response.Write "EMAIL -->" & Request.Form("Email") & "<br><br>"
'Response.Write "OPTION -->" & Request.Form("PICK1") & "<br><br>"

'----------------------------------------
'  
'----------------------------------------
F_Fullname1			= Request.Form("Fullname1")
F_Address1			= Request.Form("Address1")
F_City1				= Request.Form("City1")
F_state1			= Request.Form("state1")
F_Zipcode1			= Request.Form("Zipcode1")
F_Telephone1			= Request.Form("Telephone1")
F_Telephone_type1		= Request.Form("Telephone_type1")
F_Email1			= Request.Form("Email1")
F_Allegations1			= Request.Form("Allegations1")
F_Agency_Name1			= Request.Form("Agency_Name1")
F_Agency_POC1			= Request.Form("Agency_POC1")
F_Agency_Telephone1		= Request.Form("Agency_Telephone1")
			 		 
F_Fullname2			= Request.Form("Fullname2")
F_Address2			= Request.Form("Address2")
F_City2				= Request.Form("City2")
F_state2			= Request.Form("state2")
F_Zipcode2			= Request.Form("Zipcode2")
F_Telephone2			= Request.Form("Telephone2")
F_Telephone_type2		= Request.Form("Telephone_type2")
F_Email2			= Request.Form("Email2")
			 		 
F_Fullname3			= Request.Form("Fullname3")
F_Address3			= Request.Form("Address3")
F_City3				= Request.Form("City3")
F_state3			= Request.Form("state3")
F_Zipcode3			= Request.Form("Zipcode3")
F_Telephone3			= Request.Form("Telephone3")
F_Dis_Date_Time3		= Request.Form("Dis_Date_Time3")
			 		 
F_Race41			= Request.Form("Race41")
F_Race42			= Request.Form("Race42")
F_Race43			= Request.Form("Race43")
F_Race44			= Request.Form("Race44")
F_Race45			= Request.Form("Race45")
			 		 
F_National_Origin4		= Request.Form("National_Origin4")
F_Race4_Tribal_Affiliation4	= Request.Form("Race4_Tribal_Affiliation4")
			 		 
F_Color4			= Request.Form("Color4")
			 		 
F_Religion4			= Request.Form("Religion4")
F_Sex4				= Request.Form("Sex4")
F_Sexual_Orientation4		= Request.Form("Sexual_Orientation4")
F_Gender_Identity4		= Request.Form("Gender_Identity4")
F_Inquiring_About_Pay4		= Request.Form("Inquiring_About_Pay4")
F_Discussing_Pay4		= Request.Form("Discussing_Pay4")
F_Disclosing_Pay4		= Request.Form("Disclosing_Pay4")
F_Protected_Veteran_Status4	= Request.Form("Protected_Veteran_Status4")
F_Disability4			= Request.Form("Disability4")
F_Retaliation4			= Request.Form("Retaliation4")

F_Complaint5			= Request.Form("Complaint5")

F_Narrative5			= Request.Form("Narrative5")
F_Treated5			= Request.Form("Treated5")
F_Narrative6			= Request.Form("Narrative6")
			 		 
F_Attorney5			= Request.Form("Attorney5")
F_Fullname21			= Request.Form("Fullname21")
F_Address21			= Request.Form("Address21")
F_City21			= Request.Form("City21")
F_state21			= Request.Form("state21")
F_Zipcode21			= Request.Form("Zipcode21")
F_Telephone21			= Request.Form("Telephone21")
F_Email21			= Request.Form("Email21") 		 
			 		 
F_Print_Signature		= Request.Form("Print_Signature")
F_Print_Signature_Date		= Request.Form("Print_Signature_Date")

'----------------------------------------
'  
'----------------------------------------
Function Error_Text(txt)
	etxt = etxt & txt & "<br>"     
end function

'----------------------------------------
'  Checking input values
'----------------------------------------

if F_Fullname1 = "" then
	Error_Text("- Blank Name")
end if
	
if F_Address1 = "" then
	Error_Text("- Blank Address")
end if

if F_City1 = "" then
	Error_Text("- Blank City")
end if
	
if F_state1 = "" then
	Error_Text("- Please select a state from the dropdown")
end if

if F_Zipcode1 = "" then
	Error_Text("- Blank Zipcode")
end if
	
if F_Telephone1 = "" then
	Error_Text("- Blank Telephone #")
else
	Select Case F_Telephone_type1
  		Case 1
    		F_Telephone1 = F_Telephone1 & " (Home)"
  		Case 2
    		F_Telephone1 = F_Telephone1 & " (Work)"
  		Case 3
    		F_Telephone1 = F_Telephone1 & " (Cell)"
  		Case else
    		F_Telephone1 = F_Telephone1 & " (No selecion)"
	End Select
end if

if F_Email1 = "" then
	Error_Text("- Blank Email Address")
end if
 
if F_Allegations1 = 1 then
	txt_F_Allegations1 = "Yes"
	if F_Agency_Name1 = "" then
		Error_Text("- Blank Agency Name")
	end if
	if F_Agency_POC1 = "" then
		Error_Text("- Blank Agency Point of Contact")
	end if
	if F_Agency_Telephone1 = "" then
		Error_Text("- Blank Agency Telephone Number")
	end if
else
	txt_F_Allegations1 = "No"
	F_Agency_Name1 = "-"
	F_Agency_POC1 = "-"
	F_Agency_Telephone1 = "-"
end if

if F_Print_Signature = "" then
	Error_Text("- Please sign document")
else
	if F_Print_Signature_Date = "" then
		Error_Text("- Blank Signature Date")
	end if
end if

if F_Fullname2 <> "" then
	if F_Telephone2 = "" then
		Error_Text("- Blank Alternate Telephone #")
	else
		Select Case F_Telephone_type2
  			Case 1
    				F_Telephone2 = F_Telephone2 & " (Home)"
  			Case 2
    				F_Telephone2 = F_Telephone2 & " (Work)"
  			Case 3
    				F_Telephone2 = F_Telephone2 & " (Cell)"
  			Case else
    				F_Telephone2 = F_Telephone2 & " (No selecion)"
		End Select
	end if
end if 



if F_Race41 = "1" then
	F_Race = F_Race & "American Indian or Alaska Native Indicate Tribal Affiliation, "
end if

if F_Race42 = "1" then
	F_Race = F_Race & "Asian, "
end if

if F_Race43 = "1" then
	F_Race = F_Race & "Black or African American, "
end if

if F_Race44 = "1" then
	F_Race = F_Race & "Native Hawaiian or Other Pacific Islander, "
end if

if F_Race45 = "1" then
	F_Race = F_Race & "White"
end if


if F_Color4 = "1" then
	txt_F_Color4 	= "Yes"
else
	txt_F_Color4 	= "No"
end if

if F_Religion4 = "1" then
	txt_F_Religion4	= "Yes"
else
	txt_F_Religion4	= "No"
end if

if F_Sex4 = "1" then
	txt_F_Sex4	= "Yes"
else
	txt_F_Sex4 	= "No"
end if	 		 

if F_Sexual_Orientation4 = "1" then
	txt_F_Sexual_Orientation4	= "Yes"
else
	txt_F_Sexual_Orientation4 	= "No"
end if	 		 

if F_Gender_Identity4 = "1" then
	txt_F_Gender_Identity4	= "Yes"
else
	txt_F_Gender_Identity4 	= "No"
end if	 		 

if F_Inquiring_About_Pay4 = "1" then
	txt_F_Inquiring_About_Pay4	= "Yes"
else
	txt_F_Inquiring_About_Pay4 	= "No"
end if	 

if F_Discussing_Pay4 = "1" then
	txt_F_Discussing_Pay4	= "Yes"
else
	txt_F_Discussing_Pay4 	= "No"
end if	

if F_Disclosing_Pay4 = "1" then
	txt_F_Disclosing_Pay4	= "Yes"
else
	txt_F_Disclosing_Pay4 	= "No"
end if	

if F_Protected_Veteran_Status4 = "1" then
	txt_F_Protected_Veteran_Status4	= "Yes"
else
	txt_F_Protected_Veteran_Status4 = "No"
end if	

if F_Disability4 = "1" then
	txt_F_Disability4	= "Yes"
else
	txt_F_Disability4 	= "No"
end if

if F_Retaliation4 = "1" then
	txt_F_Retaliation4	= "Yes"
else
	txt_F_Retaliation4 	= "No"
end if



if F_National_Origin4 = "1" then
	F_National_Origin4 = "Hispanic or Latino"
else
	F_National_Origin4 = "Other"
end if

'----------------------------------------
'  if values are clean then continue
'----------------------------------------
if etxt = "" THEN

'Response.Write txt

htmlsubject = "OFCCP Complaint Form"
htmltitle = "Complaint Involving Employment Discrimination by a Federal Contractor or Subcontractor" & "<br><br>"

htmlheader = "<html><body><table cols=2 width='800px'><tr><td width='20%' align=left valign=top></td><td></td></tr>"
part1 = "<tr><th colspan=2>" & htmltitle & "<br></th></tr>"
part1 = part1 & "<tr><th colspan=2 align=left>" & "How can we reach you?" & "<br><br></th></tr>"
part1 = part1 & "<tr><td>" & "Name (First, Middle, Last):" & "</td><td>" & F_Fullname1 & "</td></tr>"
part1 = part1 & "<tr><td>" & "Adress:" & "</td><td>" & F_Address1 & "</td></tr>" 
part1 = part1 & "<tr><td>" & "City:" & "</td><td>" & F_City1 & "</td></tr>"
part1 = part1 & "<tr><td>" & "State):" & "</td><td>" & F_state1 & "</td></tr>"
part1 = part1 & "<tr><td>" & "Zipcode:" & "</td><td>" & F_Zipcode1 & "</td></tr>" 
part1 = part1 & "<tr><td>" & "Telephone:" & "</td><td>" & F_Telephone1& "</td></tr>"
part1 = part1 & "<tr><td>" & "Email:" & "</td><td>" & F_Email1 & "</td></tr>" 
part1 = part1 & "<tr><td>" & "Allegations:" & "</td><td>" & txt_F_Allegations1 & "</td></tr>"
part1 = part1 & "<tr><td>" & "Agency Name:" & "</td><td>" & F_Agency_Name1 & "</td></tr>"
part1 = part1 & "<tr><td>" & "Agenct Point Of Contact:" & "</td><td>" & F_Agency_POC1 & "</td></tr>" 
part1 = part1 & "<tr><td>" & "Agenct Telephone:" & "</td><td>" & F_Agency_Telephone1 & "</td></tr>"
part1 = part1 & "<tr><td><br><br></td><td><br><br></td></tr>"

part2 = "<tr><th colspan=2>" & "Who can we contact if we cannot reach you?" & "<br><br></th></tr>"
part2 = part2 & "<tr><td>" & "Name (First, Middle, Last):" & "</td><td>" & F_Fullname2 & "</td></tr>"
part2 = part2 & "<tr><td>" & "Adress:" & "</td><td>" & F_Address2 & "</td></tr>" 
part2 = part2 & "<tr><td>" & "City:" & "</td><td>" & F_City2 & "</td></tr>"
part2 = part2 & "<tr><td>" & "Statet):" & "</td><td>" & F_state2 & "</td></tr>"
part2 = part2 & "<tr><td>" & "Zipcode:" & "</td><td>" & F_Zipcode2 & "</td></tr>" 
part2 = part2 & "<tr><td>" & "Telephone:" & "</td><td>" & F_Telephone2 & "</td></tr>"
part2 = part2 & "<tr><td>" & "Email:" & "</td><td>" & F_Email2 & "</td></tr>" 
part2 = part2 & "<tr><td><br><br></td><td><br><br></td></tr>"

part3 = "<tr><th colspan=2>" & "What is the name of the employer that you believe discriminated or retaliated against you?" & "<br><br></th></tr>"
part3 = part3 & "<tr><td>" & "Employer Name:" & "</td><td>" & F_Fullname3 & "</td></tr>"
part3 = part3 & "<tr><td>" & "Adress:" & "</td><td>" & F_Address3 & "</td></tr>" 
part3 = part3 & "<tr><td>" & "City:" & "</td><td>" & F_City3 & "</td></tr>"
part3 = part3 & "<tr><td>" & "Statet):" & "</td><td>" & F_state3 & "</td></tr>"
part3 = part3 & "<tr><td>" & "Zipcode:" & "</td><td>" & F_Zipcode3 & "</td></tr>" 
part3 = part3 & "<tr><td>" & "Telephone:" & "</td><td>" & F_Telephone3 & "</td></tr>"
part3 = part3 & "<tr><td>" & "Email:" & "</td><td>" & F_Email3 & "</td></tr>" 
part3 = part3 & "<tr><td><br><br></td><td><br><br></td></tr>"

part4 = "<tr><th colspan=2>" & "Why do you believe your employer discriminated or retaliated against you?" & "<br><br></th></tr>"
part4 = part4 & "<tr><td>" & "Race:" & "</td><td>" & F_Race & "</td></tr>"
part4 = part4 & "<tr><td>" & "National Origin:" & "</td><td>" & F_National_Origin4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Color:" & "</td><td>" & txt_F_Color4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Religion:" & "</td><td>" & txt_F_Religion4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Sex:" & "</td><td>" & txt_F_Sex4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Sexual Orientation:" & "</td><td>" & txt_F_Sexual_Orientation4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Gender Identity:" & "</td><td>" & txt_F_Gender_Identity4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Inquiring About Pay:" & "</td><td>" & txt_F_Inquiring_About_Pay4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Discussing Pay:" & "</td><td>" & txt_F_Discussing_Pay4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Disclosing Pay:" & "</td><td>" & txt_F_Disclosing_Pay4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Protected Veteran Status:" & "</td><td>" & txt_F_Protected_Veteran_Status4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Disability:" & "</td><td>" & txt_F_Disability4 & "</td></tr>"
part4 = part4 & "<tr><td>" & "Retaliation:" & "</td><td>" & txt_F_Retaliation4 & "</td></tr>"
part4 = part4 & "<tr><td><br><br></td><td><br><br></td></tr>"

part5 = "<tr><th colspan=2>" & "Where did you learn you could file a complaint with OFCCP?" & "<br><br></th></tr>"
part5 = part5 & "<tr><td>" & "-" & "</td><td>" & F_Complaint5 & "</td></tr>"
part5 = part5 & "<tr><td><br><br></td><td><br><br></td></tr>"

part6 = "<tr><th colspan=2>" & "Your Complaint: <br>Please describe below what you think the employer did or did not do that you believe caused discrimination or retaliation, including: " & "<br><br></th></tr>"
part6 = part6 & "<tr><td>" & "Description:" & "</td><td>" & F_Narrative5 & "</td></tr>"
part6 = part6 & "<tr><td><br><br></td><td><br><br></td></tr>"

part7 = "<tr><th colspan=2>" & "Do you think the discrimination includes or affects others?" & "</th></tr>"
part7 = part7 & "<tr><td>" & "Were employees or applicants treated the same?" & "</td><td>" & F_Treated5 & "</td></tr>"
part7 = part7 & "<tr><td>" & "Description" & "</td><td>" & F_Narrative6 & "</td></tr>"
part7 = part7 & "<tr><td><br><br></td><td><br><br></td></tr>"

part8 = "<tr><th colspan=2>" & "Do you have an attorney or other representative?" & "<br><br></th></tr>"
part8 = part8 & "<tr><td>" & "Attorney/Representatives Name:" & "</td><td>" & F_Fullname21 & "</td></tr>"
part8 = part8 & "<tr><td>" & "Adress:" & "</td><td>" & F_Address21 & "</td></tr>" 
part8 = part8 & "<tr><td>" & "City:" & "</td><td>" & F_City21 & "</td></tr>"
part8 = part8 & "<tr><td>" & "Statet):" & "</td><td>" & F_state21 & "</td></tr>"
part8 = part8 & "<tr><td>" & "Zipcode:" & "</td><td>" & F_Zipcode21 & "</td></tr>" 
part8 = part8 & "<tr><td>" & "Telephone:" & "</td><td>" & F_Telephone21 & "</td></tr>"
part8 = part8 & "<tr><td>" & "Email:" & "</td><td>" & F_Email21 & "</td></tr>" 
part8 = part8 & "<tr><td>" & "Who should we contact if we need more information about your description of what occurred?" & "</td><td>" & F_Attorney5 & "</td></tr>"
part8 = part8 & "<tr><td><br><br></td><td><br><br></td></tr>"

part9 = "<tr><th colspan=2>" & "Signature and Verification" & "</th></tr>"
part9 = part9 & "<tr><td>" & F_Print_Signature & "</td><td>" & F_Print_Signature_Date & "</td></tr>"
part9 = part9 & "<tr><td></td><td></td></tr>" 

htmltail = "</table></body></html>"


'----------------------------------------
' Send email to a specific region based on State
'----------------------------------------
	
	Select Case F_state1
		'---------------------------------------------
		'These states are assigned to Region OFCCP-NE-CC4@DOL.GOV
		'New Jersey, New York, Puerto Rico,Virgin Islands, Connecticut, Maine, Massachusetts, 
		'New Hampshire, Rhode Island or Vermont
		'---------------------------------------------
		case "NJ" , "NY" , "PR" , "VI" , "CT" , "ME" , "MA" , "NH" , "RI" , "VT" 
			region = 1
			region_title 	= ""
			region_email	= "OFCCP-NE-CC4@DOL.GOV"

		'---------------------------------------------
		'These states are assigned to Region OFCCP-MA-CC4@DOL.GOV
		'Delaware, District of Columbia, Maryland, Pennsylvania, Virginia, West Virginia
		'---------------------------------------------
		case "DE" , "DC" , "MD" , "PA" , "VA" , "WV" 
			region = 2
			region_title 	= ""
			region_email	= "OFCCP-MA-CC4@DOL.GOV"

		'---------------------------------------------
		'These states are assigned to Region OFCCP-SE-CC4@DOL.GOV
		'Alabama, Florida, Georgia, Kentucky, Mississippi, North Carolina, South Carolina, Tennessee
		'---------------------------------------------
		case "AL" , "FL" , "GA" , "KY" , "MS" , "NC" , "SC" , "TN"
			region = 3
			region_title 	= ""
			region_email	= "OFCCP-SE-CC4@DOL.GOV"

		'---------------------------------------------
		'These states are assigned to Region OFCCP-MW-CC4@DOL.GOV
		'Illinois, Indiana, Iowa, Kansas, Michigan, Minnesota, Missouri, Nebraska, Ohio, Wisconsin
		'---------------------------------------------
		case "IL" , "IN" , "LA" , "KS" , "MI" , "MN" , "MO", "NE" , "OH" , "WI"
			region = 4
			region_title 	= ""
			region_email	= "OFCCP-MW-CC4@DOL.GOV"

		'---------------------------------------------
		'These states are assigned to Region OFCCP-SW-CC4@DOL.GOV
		'Arkansas, Colorado, Louisiana, Montana, New Mexico, North Dakota, Oklahoma, 
		'South Dakota, Texas, Utah, Wyoming
		'---------------------------------------------
		case "AR" , "CO" , "MT" , "NM" , "ND" , "OK" , "SD", "TX" , "UT" , "WY"
			region = 5
			region_title 	= ""
			region_email	= "OFCCP-SW-CC4@DOL.GOV"

		'---------------------------------------------
		'These states are assigned to Region OFCCP-PA-CC4@DOL.GOV
		'Alaska, Arizona, California, Guam, Hawaii, Idaho, Nevada, Oregon, Washington
		'---------------------------------------------
		case "AK" , "AZ" , "CA" , "GU" , "HI" , "ID" , "NV", "OR" , "WA"
			region = 6
			region_title 	= ""
			region_email	= "OFCCP-PA-CC4@DOL.GOV"

	End Select


		html = ""
		knt = 0
		txt = ""

		'----------------------------------------
		'  SEND TO REGION
		'----------------------------------------
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = "OFCCP - Department Of Labor"
		Mailer.FromAddress= mail_to
		Mailer.RemoteHost = "smtp.rcn.com"
		'Mailer.AddRecipient region_email, "ddunston@viateq.com"
		Mailer.AddRecipient region_email, region_email
		Mailer.Subject    = htmlsubject  
		Mailer.BodyText   = htmlheader & part1 & part2 & part3 & part4 & part5 & part6 & part7 & part8 & part9 & htmltail
		Mailer.ContentType = "text/html"

		if Mailer.SendMail then
		  'Response.Write "Mail sent...<br><br>"
		else
		  'Response.Write "Mail send failure. Error was " & Mailer.Response
		end if

		'----------------------------------------
		' SEND TO COMPLAINANT
		'----------------------------------------
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = "OFCCP - Department Of Labor"
		Mailer.FromAddress= mail_to
		Mailer.RemoteHost = "smtp.rcn.com"
		Mailer.AddRecipient F_Fullname1, F_Email1
		Mailer.Subject    = htmlsubject  
		Mailer.BodyText   = htmlheader & part1 & part2 & part3 & part4 & part5 & part6 & part7 & part8 & part9 & htmltail
		Mailer.ContentType = "text/html"

		if Mailer.SendMail then
		  'Response.Write "Mail sent...<br><br>"
		else
		  'Response.Write "Mail send failure. Error was " & Mailer.Response
		end if



		
		response.redirect "english-form.html"
else 
		Response.Write "<p>Please click Back-Space on your browser and review each entry:</p>"
		Response.Write "COMPLAINT FORM ERROR<br>--------------------------<br><br>" & etxt
End if
%>