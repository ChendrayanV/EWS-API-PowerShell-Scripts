function New-xAppointment {
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
        $ExchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
        $ExchangeService.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        $ExchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        $Appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::new($ExchangeService)
        $Appointment.Subject = "Demo Appointment"
        $Appointment.Start = [datetime]::Now.AddDays(2)
        $Appointment.End = $Appointment.Start.AddHours(1)
        $Appointment.Location = "Building 1 - Reception Area"
        $Appointment.ReminderDueBy = [datetime]::Now.AddMinutes(1)
        $Appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
    }

    end {

    }
}