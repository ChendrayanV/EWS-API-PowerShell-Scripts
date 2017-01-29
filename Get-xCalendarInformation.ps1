function Get-xCalendarInformation {
    [outputtype('Microsoft.Exchange.WebServices.Data.Item')]
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
        $ExchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
        $ExchangeService.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        $ExchangeService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailBox)
        $ExchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        $CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($ExchangeService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        $View = [Microsoft.Exchange.WebServices.Data.CalendarView]::new([datetime]::Now,[datetime]::Now.AddDays(2))
        $CalendarFolder.FindAppointments($View)
    }

    end {

    }
}