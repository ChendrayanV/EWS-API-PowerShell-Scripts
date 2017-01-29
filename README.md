# Using EWS Managed API 2.2 in PowerShell
### Summary 
This repository is created to share the PowerShell scripts used for a demo in Bangalore PowerShell User Group (January, 21 2017)
- Get-xTask.ps1
- Get-xCalendarInformation.ps1
- New-xTask.ps1
- New-xMeeting.ps1
- New-xAppointment.ps1
- Get-xPhoneAlert.ps1

### Disclaimer
- We haven't included exception handling in the scripts. 
- All the script is built for learning and demo purpose.
- Scripts are built __on the fly__ without any testing. 

# Get Started
To get started we need EWS Managed API (EwsManagedApi.msi) package and here is the [Download Link](http://www.microsoft.com/en-in/download/details.aspx?id=42951) and the installation is just 
Click -> Click -> Click. Once the installation is completed successfully we find the assemblies in the below path 
> ```"C:\Program Files\Microsoft\Exchange\Web Services\2.2"```. 
>
Below are the files get installed by the EwsManagedApi.msi
- Microsoft.Exchange.WebServices.dll   The signed .NET assembly that implements the EWS Managed API.
- Microsoft.Exchange.WebServices.xml   Provides a Visual Studio .NET IntelliSense file for the EWS Managed API.
- Microsoft.Exchange.WebServices.Auth.dll   Provides an API to validate, parse and process Exchange Identity Tokens to be used by mail apps for Outlook.
- Microsoft.Exchange.WebServices.Auth.xml   Provides a Visual Studio .NET IntelliSense file for the Auth library.
- GettingStarted.doc   Provides additional help for and information about how to use the API.
- License Terms.rtf   Contains the license terms for using the EWS Managed API and documentation.
- Readme.htm   This file (Get Started File - Has more links for our reference)
- Redist.txt   Defines which files and directories can be redistributed under the license terms.

# Get-xTask
A PowerShell script to list tasks from the given mailbox.

### Example 1
> ```Get-xTask -MailBox user1@tenant.onmicrosoft.com -ItemCount 5 -Credential admin@tenant.onmicrosoft.com```
>
> List top 5 task from the user user1@tenant.onmicrosoft.com

### Example 2
> ```"user1@tenant.onmicrosoft.com" , "user2@tenant.onmicrosoft.com" | Get-xTask -ItemCount 5 -Credential admin@tenant.onmicrosoft.com```
>
> List top 5 task from the users (user1 and user2)

### Example 3
> ```"user1@tenant.onmicrosoft.com" , "user2@tenant.onmicrosoft.com" | Get-xTask -ItemCount 5 -Credential admin@tenant.onmicrosoft.com | ? {$_.Status -eq 'InProgress'}```
>
> List top 5 task from the users (user1 and user2) where the status is in progress status

### Example 4 
> ```'user2@tenant.onmicrosoft.com' , 'user1@tenant.onmicrosoft.com' | Get-xTask -ItemCount 5 -Credential user1@tenant.onmicrosoft.com  | Select Owner , Status , StartDate ```
>
> List top 5 task from the users (user1 and user2) select the properties required 

# Get-xCalendarInformation
A PowerShell script to list calendar information (Includes appointments and meetings).

### Example 1
> ```Get-xCalendarInformation -MailBox user1@tenant.onmicrosoft.com -Days 3 -Credential user1@tenant.onmicrosoft.com```
>
> List calendar information from the given mailbox 

### Example 2
> ```Get-xCalendarInformation -MailBox user1@tenant.onmicrosoft.com -Days 3 -Credential user1@tenant.onmicrosoft.com | Select Subject , Importance```
>
> List the property required like subject, importance etc. 

### Example 3
> ```'user1@tenant.onmicrosoft.com' , 'user2@tenant.onmicrosoft.com' | Get-xCalendarInformation -Credential user1@tenant.onmicrosoft.com | Select Subject , Importance```
>
> List the property required like subject, importance etc for more than one user

# New-xTask 
A PowerShell script to create a task for the user(s)

### Example 1 
> ```New-xTask -Mailbox user1@tenant.onmicrosoft.com -Credential user1@tenant.onmicrosoft.com```
>
> Creates a new task with a subject 'New Task 1' and body as 'Check This!' (Inlined in the script)

### Example 2
> ```'user1@tenant.onmicrosoft.com' , 'user2@tenant.onmicrosoft.com' | New-xTask -Credential user1@tenant.onmicrosoft.com```
>
> Creates a new task with a subject 'New Task 1' and body as 'Check This!' (Inlined in the script) for more than one user. 

# New-xMeeting 
A PowerShell script to create a new meeting request. Few properties like ```$Start```, ```$End``` and ```$Subject``` are inlined. 
Modify it or parameterize it as required. For example, we can change start and end date here! 
>```$Meeting.Start = [datetime]::Now.AddDays(2)```
>
>```$Meeting.End = $Meeting.Start.AddHours(1)```
>
Note: Apply the logic for date and time using ```[datetime]```
### Example 1
>```New-xMeeting -MailBox user1@tenant.onmicrosoft.com -Credential user1@tenant.onmicrosoft.com```
>
>Creates a meeting request in which one is optionals and other in required. 

# New-xAppointment
A PowerShell script to create an appointment. Meeting needs a recipients address whereas appointment doesn't. 
Instantiated the Appointment class and now the $Appointment is an object.
> ```$Appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::new($ExchangeService)```
>
Set the subject as needed or parameterize it.
> ```$Appointment.Subject = "Demo Appointment"```
>
In our case the start date is 2 days from now! 
>
> ```$Appointment.Start = [datetime]::Now.AddDays(2)```
>
The appointment is booked for an hour.
>
> ```$Appointment.End = $Appointment.Start.AddHours(1)```
>
Define the location.
>
> ```$Appointment.Location = "Building 1 - Reception Area"```

### Example 1
> ```New-xAppointment -Mailbox user1@tenant.onmicrosoft.com -Credential user1@tenant.onmicrosoft.com ```
>
> Books an appointment for the given mail box (For one mailbox)

### Example 2
> ```'user1@tenant.onmicrosoft.com' , 'user1@tenant.onmicrosoft.com' | New-xAppointment -Credential 'user1@tenant.onmicrosoft.com'```
>
> Books an appointment for more than one user with same information. 

# Get-xPhoneAlert
A customer asked me a solution to play a phone call when there is a Important email. Yes, we need to do Event Subscriptions but for a 
demo we shared a piece of PS Script which scans inbox and dials the given telephone to and play the text. Below is the logic and ensure 
you have UM (Unified Messaging) enabled mailbox! 
Get the ItemID of the mail which is required for the PlayOnPhoneMethod 
> ```$ItemId = $Results.Items[0].Id.UniqueId```
>
Invoke the method UnifiedMessaging and PlayOnPhone with two overloads (ItemId and PhoneNumber)
> ```$Call = $ExchangeService.UnifiedMessaging.PlayOnPhone($ItemId, <+CountryCode><Number>)```
>

### Example 1
> ```Get-xPhoneAlert -MailBox chendrayan.venkatesan@contoso.com -PhoneNumber +91<Number> -Credential chendrayan.venkatesan@contoso```
>
> Scan the inbox of chendrayan.venkatesan@contoso.com to retrieve High priority email and play the text on phone 

# References
- [EWS API 2.2 Download Link](http://www.microsoft.com/en-in/download/details.aspx?id=42951)
- [Get started with EWS Managed API client applications](https://msdn.microsoft.com/en-us/library/office/dn567668(v=exchg.150).aspx)
- [How to communicate with EWS by using the EWS Managed API](https://msdn.microsoft.com/en-us/library/office/dn467891(v=exchg.150).aspx)
