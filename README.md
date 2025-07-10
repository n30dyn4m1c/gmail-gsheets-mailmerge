# Gmail and Google Sheets Mail Merge using AppScript
Scripts and steps for mail merge using Gmail and Google Sheets 

## Setup:

Step 1: Upload Folders with PDFs of Offer Letters and Invoices to Google Drive

Step 2: Create Google Sheet and name it GetFileID.

Step 3: Go to GetFileID Google Sheet. Remove all data from sheets except for columns with formulae.

Step 4: While Sheet 1 is active, go to Extensions->Apps Script

Step 5: Copy & paste getfileid.js code.

Step 6: Replace Google Drive folder ID in the code.

Step 7: Click Run Code to run the function. This will save the File Names and File IDs in Sheet 1

Step 8: Paste the File Name and File ID data in the next tab Formatted Data, then you should see key parts separated in the columns with formulae.
- File Name – Paste from Sheet 1
- File ID – Paste from Sheet 1
- ID – =LEFT(A2,9)
- StudentID – =SUBSTITUTE(C2,”_”,””)
- FileName – =A2
- FileID – =B2

Step 9: Apply a filter on the sheet Formatted Data.
- Filter by condition, text contains: “OfferLetter”
- Paste result into FileID_OfferLetter tab

- Filter by condition, text contains: “Invoice”
- Paste result into FileID_Invoice tab

Step 10: Open Excel sheet with StudentID, Name, Surname, Email columns and copy these
- Paste into CompiledData sheet.

- Student ID – Paste from Excel student details with email sheet
- Name – Paste from Excel student details with email sheet
- Surname – Paste from Excel student details with email sheet
- Email – Paste from Excel student details with email sheet
- OfferLetterFileID – =VLOOKUP(A2,FileID_OfferLetter!$A$1:$C$1000,3, FALSE)
- OfferLetterFileName – =VLOOKUP(A2,FileID_OfferLetter!$A$1:$C$1000,2, FALSE)
- InvoiceFileID – =VLOOKUP(A2,FileID_Invoice!$A$1:$C$1000,3, FALSE)
- InvoiceFileName – =VLOOKUP(A2,FileID_Invoice!$A$1:$C$1000,2, FALSE)
- Check – =ISNUMBER(SEARCH(C2,F2))

Step 11: Ensure the Check column has TRUE for all. If some are showing #NA, then go through the steps again to double check.
- Count the number of records while going through to steps see that they all add up.

Step 12: Copy all data in the CompiledData sheet except for the Check Column.

Step 13: Open the file 2025_Offer_Invoice_MailMerge and paste data into next blank columns in Sheet1

- If doing for next year, make a copy of this file and rename, but use the same function saved in AppsScript.

Step 14: Double Check file IDs against names. The next tab has the url format to paste and double check https://drive.google.com/file/d/YOUR_FILE_ID_HERE/view

- Do random checks on at least 5 records.

Step 15: Paste the code for mail-merge.js into Extensions->Apps Script

Step 16: Prepare a Draft Email with appropriate subject and content

Subject:
- UPNG Offer Letter and Invoice for 2025

Content:
Dear {{FirstName}} {{LastName}},

We are delighted to inform you that you have been accepted to study at the University of Papua New Guinea in 2025.

Attached, you will find your offer letter and invoice. Please review the offer letter for important details and follow the provided instructions to complete your enrollment.

Best regards,

Roboam Kakap
Acting Registrar
University of Papua New Guinea
PO Box 320, University PO,
National Capital District
Papua New Guinea

Note that this is an automated email. Please do not reply to this email.

Step 17: If you want to test, paste 5 records, then replace emails with your personal email address
- When ready to send, click Mail Merge -> Send Emails

It will ask for the Subject Line of the draft email. Paste it: 
- UPNG Offer Letter and Invoice for 2025

Step 18: If test is successful, paste the final data, check random files again, then again click Mail Merge -> Send Emails, Paste Email Subject

Step 19: Wait for script to run and send emails, also wait for failed email notifications if you wish to record failed emails

Other Options:

- Consider if you want to cc another work email
- Check the content body and update
- Make sure the subject is correct
- Consider if you want attach other files or link to them
