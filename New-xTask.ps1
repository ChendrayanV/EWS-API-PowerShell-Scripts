function New-xTask {
    [Outputtype('Microsoft.Exchange.WebServices.Data.Task')]
    param 
    (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        $Mailbox,

        [Parameter(Mandatory)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]
        $Credential
    )

    begin {
        Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    }
    
    process {
        $ExchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)
        $ExchangeService.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        $ExchangeService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$Mailbox)
        $ExchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        $Task = [Microsoft.Exchange.WebServices.Data.Task]::new($ExchangeService)
        $Task.Subject = "New Task 1"
        $Task.Body = [Microsoft.Exchange.WebServices.Data.MessageBody]::new([Microsoft.Exchange.WebServices.Data.BodyType]::Text,"Check this!")
        $Task.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Tasks)
    }

    end {

    }
}
