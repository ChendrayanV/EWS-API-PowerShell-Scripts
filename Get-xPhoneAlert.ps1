function Get-xPhoneAlert {
    param 
    (
        [Parameter(Mandatory)]
        $MailBox,

        # Include Country Code E.G (+31 for Nederlands or +91 for India)
        [Parameter(Mandatory)]
        $PhoneNumber,

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

        $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
        $View = [Microsoft.Exchange.WebServices.Data.ItemView]::new(1)
        $SearchFilter = [Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo]::new([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Importance, 
            "High")
        $Results = $ExchangeService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$SearchFilter,$View)
        $ItemId = $Results.Items[0].Id.UniqueId
        $Call = $ExchangeService.UnifiedMessaging.PlayOnPhone($ItemId, $PhoneNumber)
        $Call.Refresh()
    }

    end {
    }
}