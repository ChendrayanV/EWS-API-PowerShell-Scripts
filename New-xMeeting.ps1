function New-xMeeting {
    param 
    (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        $MailBox,

        [Parameter(Mandatory)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]
        $Credential
    )

    begin {
        Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    }

    process {
        $ExchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
        $ExchangeService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailBox)
        $ExchangeService.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        $ExchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        $Meeting = [Microsoft.Exchange.WebServices.Data.Appointment]::new($ExchangeService)
        $Meeting.Start = [datetime]::Now.AddDays(2)
        $Meeting.End = $Meeting.Start.AddHours(1)
        $Meeting.Subject = "Managed EWS API Demo in PowerShell"
        $Meeting.RequiredAttendees.Add('karthik@ChensOffice365.onmicrosoft.com')
        $Meeting.OptionalAttendees.Add('MostafaSelim@ChensOffice365.onmicrosoft.com')
        $Meeting.ReminderMinutesBeforeStart = 15;
        $Meeting.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
    }

    end {

    }
}