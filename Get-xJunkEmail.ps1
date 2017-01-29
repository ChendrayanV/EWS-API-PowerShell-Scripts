function Get-xJunkEmail {
    [outputtype('Microsoft.Exchange.WebServices.Data.EmailMessage')]
    param 
    (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        $MailBox,

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
        $ExchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
        $ExchangeService.ImpersonatedUserId =[Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailBox)
        $ExchangeService.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        $ExchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        if($PSBoundParameters.ContainsKey('ItemCount')) {
            $View = [Microsoft.Exchange.WebServices.Data.ItemView]::new($ItemCount)
        }
        else {
            $View = [Microsoft.Exchange.WebServices.Data.ItemView]::new(10)
        }
        $ExchangeService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail,$View)
    }

    end {

    }
}