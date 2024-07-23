/************************************************************************************
***** Program: 	Email Notification Macro *****
***** Author:	joshkylepearce           *****
************************************************************************************/

/************************************************************************************
Distribution List Setup

Purpose:
Define the distribution list that should be sent results via email.
The recipients of a process are subject to change. 
Handling the distribution list outside of the macro enables users to easily change
recipients without the requirement to update the email notification macro.

Format:
To ensure compatibility with the macro, the distribution list must be:
-	Saved as 'distribution_list'.
-	Contain two variables named TO_CC_BCC & EMAIL_ADDRESS.
-	Additional recipients can be emailed by adding a new row.
-	Email addresses can contain up to 254 characters.
-	The recipient options are TO, CC, & BCC.
-	TO, CC, & BCC are not case sensitive.

Format Example:
TO_CC_BCC	EMAIL_ADDRESS
TO			joshkylepeare@gmail.com
TO			user1@github.com
TO			user2@github.com
CC			user3@github.com
BCC			user4@github.com

Code Example:
data distribution_list;
infile datalines delimiter=",";
format TO_CC_BCC $3. EMAIL_ADDRESS $254.;
input TO_CC_BCC $ EMAIL_ADDRESS $;
datalines;
TO, joshkylepearce@gmail.com
TO,	user1@github.com
TO,	user2@github.com
CC,	user3@github.com
BCC,user4@github.com
run;
************************************************************************************/

/*Define the distribution list to be emailed*/
data distribution_list;
infile datalines delimiter=",";
format TO_CC_BCC $3. EMAIL_ADDRESS $254.;
input TO_CC_BCC $ EMAIL_ADDRESS $;
datalines;
TO, joshkylepearce@gmail.com
run;

/************************************************************************************
Email Notification Macro

Purpose:
Sends an email notification to the distribution list of interest.

Input Parameters:
1.	user_email	- Email address of that the email will be sent from.
2. 	user_name	- Name of sender to be included in the email signature.
3.	subject		- Text contained in the email subject.
4.	mail_text	- Text contained in the email body.
5.	attach		- File(s) attached to the email.

Output Parameters:
1. 	to_email	- Email address(es) of recipient(s).
2.	cc_email	- Email address(es) of carbon copy recipient(s)
3. 	bcc_email	- Email address(es) of blind carbon copy recipient(s).
4. 	from_email	- Email address of sender. 

Macro Usage:
1. 	Create a distribution list. See section above for details.
2. 	Run the email_notification macro code.
3. 	Call the email_notification macro and enter the input parameters.
	e.g. 
	%email_notification(
	user_email	= "joshkylepearce@gmail.com",
	user_name	= "Josh Pearce",
	subject		= "REPORT",
	mail_text	= "Please find attached the latest report.",
	attach		= "\\sasebi\USER\joshkylepearce\output_file.xlsx"
	);

Notes:
-	Ensure that all input parameters are contained in quotes.
-	If no attachment is required, leave attach as blank (no quotations required).
-	Macro variables enhance attach parameter usage e.g. 
	attach="&outpath.\&outfile._&rpt_date..xlsx"
************************************************************************************/

%macro email_notification(user_email,user_name,subject, mail_text, attach);

/*Define email address used to send results to stakeholders*/
%let from_email = &user_email.;

/*Initialize email recipient macros*/
%let to_email = ;
%let cc_email = ;
%let bcc_email = ;

/*Determine whether the email recipient should be addressed to, cc'd or bcc'd*/
data _null_;
	set Distribution_List;
	EMAIL_ADDRESS=quote(trim(EMAIL_ADDRESS),"'");
		/*Determine which stakeholders should be emailed*/
		if upcase(TO_CC_BCC) = "TO" then
			call symput('to_email',catx(' ',symget('to_email'),EMAIL_ADDRESS));
		/*Determine which staleholders should be cc'd*/
		else if upcase(TO_CC_BCC) = 'CC' then
			call symput('cc_email',catx(' ',symget('cc_email'),EMAIL_ADDRESS));
		/*Determine which stakeholders should be bcc'd*/
		else if upcase(TO_CC_BCC) = 'BCC' then 
			call symput('bcc_email',catx(' ',symget('bcc_email'),EMAIL_ADDRESS));
run;

/*Write the 'to', 'cc', 'bcc' & 'from' list to the SAS log*/
%put &to_email;
%put &cc_email;
%put &bcc_email;
%put &from_email.;

/*Create conditional logic to account for emails with/without attachments*/
/*If attach input parameter is not listed, do not include attachment*/
%if %length(&attach.) = 0 %then %do;
	/*Email setup*/
	filename outbox email 
	/*State email sender*/
	from = (&from_email.)
	/*State email recipients based on the imported distribution list*/
	to = (&to_email.)
	cc = (&cc_email.)
	bcc = (&bcc_email.)
	/*State the subject of the email notification*/
	subject = &subject.
	;
%end;
/*If attach input parameter is listed, include attachment*/
%else %do;
	/*Email setup*/
	filename outbox email 
	/*State email sender*/
	from = (&from_email.)
	/*State email recipients based on the imported distribution list*/
	to = (&to_email.)
	cc = (&cc_email.)
	bcc = (&bcc_email.)
	/*State the subject of the email notification*/
	subject = &subject.
	/*Attach output to the email notification*/
	attach = &attach.
	;
%end;

/*Send an email notification to relevant stakeholders*/
data _null_;
	file outbox;
	/*Write the main text to the body of the email*/
	put "Hi,";
	put ;
	put &mail_text.;
	put ;
	put "Best regards,";
	put &user_name.;
run;

%mend;

/************************************************************************************
Example 1: SASHELP.TOURISM
************************************************************************************/

/************************************************************************************
Example 1: Data Setup
************************************************************************************/

/*Define file explorer location of output*/
%let outpath=\\sasebi\SAS User Data\Josh Pearce\DATA;

/*Define the filename of the data to be exported*/
%let report_name1=TOURISM;

/*Define the file explorer name & extension*/
%let outfile1="&outpath.\&report_name1..xlsx";

/*Export dataset to file explorer as a xlsx file*/
proc export 
data=SASHELP.TOURISM
outfile=&outfile1.
dbms=xlsx
replace;
run;

/************************************************************************************
Example 1: Macro Usage
************************************************************************************/

/*Email distribution list based on input parameters*/
%email_notification(
user_email	= "joshkylepearce@gmail.com",
user_name	= "Josh Pearce",
subject		= "TOURISM",
mail_text	= "Please find attached the SASHELP.TOURISM table.",
attach		= &outfile1.
);

/************************************************************************************
Example 2: SASHELP.RENT
************************************************************************************/

/************************************************************************************
Example 2: Data Setup
************************************************************************************/

/*Define the filename of the data to be exported*/
%let report_name2=RENT;

/*Create macro for the date (format MMMYY) listed in the report output filename*/
%let report_month = %sysfunc(intnx(month,"&sysdate."d,-1),monyy7.);
%put &report_month;

/*Define the file explorer name & extension*/
%let outfile2="&outpath.\&report_name2._&report_month..xlsx";

/*Export dataset to file explorer as a xlsx file*/
proc export 
data=SASHELP.RENT
outfile=&outfile2.
dbms=xlsx
replace;
run;

/************************************************************************************
Example 2: Macro Usage
************************************************************************************/

/*Email distribution list based on input parameters*/
%email_notification(
user_email	= "joshkylepearce@gmail.com",
user_name	= "Josh Pearce",
subject		= "&report_name2. &report_month.",
mail_text	= "Please find attached the latest &report_name2. report.",
attach		= &outfile2.
);