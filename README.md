# Outlook Addin: Recurring E-mail Messages # 
*This is an Outlook addin for setting up recurring e-mail messages by using calendar appointments. 
Dealing with dynamic dates is possible with embedded date variables.*

As a finance professional I have developed and been using a modified version of this addin for some time in my day to day tasks.  This is a generic version for others' use regardless of the field of profession performed. 

Addin is offering four main functions.  

## 1. Sending plain e-mail messages without attachments

Function Name: 
Send Recurring Email

Description:
This function sends plain e-mails without any attachments.  

How to use:
-	You should enter your e-mail body text into appointment’s location field.  
-	You should enter your recipients into appointment’s body field
-	You should enter your e-mail subject into appointment’s title field
-	Select (or first create) category named “SendRecurringEmail”.
-	Turn on reminder.  Set recurrence frequency and all done.
-	When the appointment’s reminder pops out macro/function will run.
-	You should dismiss the reminder after it pops out to avoid possible duplicate e-mail submissions. 

Sample appointment entry:

![image](https://user-images.githubusercontent.com/59412630/178556239-4d1cbcd3-661c-45e3-b6cf-8ee256cc3b2d.png)


## 2. Sending e-mail messages with attachments
Function Name: 
Send Recurring Email with Attachment

Description:
This function tries to send the requested file to your recipients.  
If file does not exist or is unreachable, function sends an e-mail saying file not found.  (If you want to use this function to send some formal e-mail messages with files to people outside your company, you should create a similar function which will not send “the file not found” message to all recipients.  Sending this message only to organizer will be more appropriate.) 

How to use:
-	You should enter your attachment’s path & file into appointment’s location field.  
-	You should enter your recipients into appointment’s body field
-	You should enter your e-mail subject into appointment’s title field.
-	Create a new category in Outlook with name “SendFile” (if you have not created before) and select this category.
-	Turn on reminder.  Set recurrence frequency.  All done.
-	When the appointment’s reminder pops out function will run and try to find and send your file to your recipients.
-	You should dismiss the reminder after it pops out to avoid possible duplicate e-mail submissions. 

Sample appointment entry:

![image](https://user-images.githubusercontent.com/59412630/178555866-86e45aab-ce97-4ef7-95c3-d4e2086ce633.png)

## 3. Checking proofs for critical tasks and informing team members accordingly
Function Name: 
Check Proof File

Description:
There might be some tasks need to be completed by cut-off times throughout a day.  Missing any of these cut-offs could lead some troubles with counterparties or penalties imposed by regulators.  Therefore, saving a proof (a screenshot in jpg format) after completion provides kind of an insurance against possible incidents.  
“Check Proof File” function looks for these jpg files at specific times of day.  If it does not find the proof file, sends an e-mail to recipients (team members in most cases) with a text message of “… file not found, task may not be completed”. 
If it finds the file, attaches it to an e-mail and sends to recipients with a text message of “… task is completed.  Click to see proof / screenshot.” 

How to use:
-	You should enter path of proof file into appointment’s location field.  
-	You should enter file name (with suffix like [tax.jpg]) into appointment’s title field.
-	You should enter your recipients into appointment’s body field
-	Create a new category in Outlook with name “CheckFileProof” (if you have not created before) and select this category.
-	Turn on reminder.  Set recurrence frequency.  All done.
- When the appointment’s reminder pops out macro/function will run.
-	You should dismiss the reminder after it pops out to avoid possible duplicate e-mail submissions. 

Sample appointment entry:

![image](https://user-images.githubusercontent.com/59412630/178557948-40af6492-79b2-48b4-b397-6c8258385135.png)

## 4. Checking proofs and informing only absence
Function Name: 
Check Proof and Inform Only Absence

Description:
This function is a slightly different version of Check Proof File function (see related above explanation first).  While “check proof file” function informs about both the existence and non-existence of target file, “Check Proof and Inform Only Absence” function only informs the recipients if target file is not available.  You can prefer using this function if proofs need to be checked more then one time throughout the day.  You will not be informed more than once in a day if proof checking is successful.  

How to use:
-	You should enter path of proof file into appointment’s location field.  
-	You should enter file name (with suffix like [tax.jpg]) into appointment’s title field.
-	You should enter your recipients into appointment’s body field
-	Create a new category in Outlook with name “CheckProofAndInformOnlyAbsence” (if you have not created before) and select this category.
-	Turn on reminder.  Set recurrence frequency.  All done.
-	When the appointment’s reminder pops out macro/function will run.
-	You should dismiss the reminder after it pops out to avoid possible duplicate e-mail submissions. 

Sample appointment entry:

![image](https://user-images.githubusercontent.com/59412630/178557812-ca0607c2-f377-4a1c-bae6-30625c3282d5.png)

## Variables available

Following variables can be used while creating recurring items.  Most of them are used to define dates dynamically (If you need new ones, you can embed them into function “EvaluateFolder”).

Variables should be entered in square brackets.  See samples below.

![image](https://user-images.githubusercontent.com/59412630/178559920-07ce832c-ae26-4c85-9f66-580ad83607c4.png)	

Sample appointment entry:

**Below entry will return “C:\Reports\202206\Daily\30\” on July, 1st.**

![image](https://user-images.githubusercontent.com/59412630/178559467-adfdb6af-cddb-46e4-9fab-a00c598f874a.png)

 
## Other Remarks
**Before installing, you should create file "appSettings.xml" in folder 'c:\temp' and save below lines in it.**
**Addin will also create empty txt files in C:\temp folder to determine if a task is run before to avoid duplicate submissions in some cases.** 
  ```
<appSettings>
    <add key="holidays" value="29.10.2022, 30.08.2022, 15.07.2022, 12.07.2022, 11.07.2022, 19.05.2022, 04.05.2022, 03.05.2022, 02.05.2022" />
  </appSettings>
  ```
**These are the non-working days in current year and needed to determine correct reporting dates.**  
**There is also a button on the addin's ribbon to open and update this xml file anytime.**

## Addin's ribbon
Below ribbon will show up after installation.

![image](https://user-images.githubusercontent.com/59412630/178563039-189b42d4-2039-4256-a07b-80a7d54849fa.png)

- If you want to pause addin's tasks for some time, you can turn it off from ribbon.
- Above given documentation is also available from ribbon's function help tab.
- User-defined macros can be added to addin and easyly be called from macro buttons in ribbon.



