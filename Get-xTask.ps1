function Get-xTask {
    [Outputtype('Microsoft.Exchange.WebServices.Data.Task')]
    param 
    (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        $Mailbox,

        [Parameter()]
        $ItemCount,

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
        if($PSBoundParameters.ContainsKey('ItemCount')) {
            $View = [Microsoft.Exchange.WebServices.Data.ItemView]::new($ItemCount)
        }
        else {
            $View = [Microsoft.Exchange.WebServices.Data.ItemView]::new(10)
        }
        $ExchangeService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Tasks,$View)
    }

    end {

    }
}
