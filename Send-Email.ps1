# Get the credential
$credential = Get-Credential

##CSV string seperated by ; character
$To = ''

## Parameters for email content
$mailSpec = @{
    SmtpServer                 = 'smtp.office365.com'
    Port                       = '587'
    UseSSL                     =  $true 
    Credential                 =  $credential
    From                       = ''
    To                         = ''
    BCC                        =  $To.Split(';')
    Subject                    = 'BAU Refresh - provision of new Dell laptop to replace current laptop - setting of temporary password - URGENT'
    Priority                   = 'high'
    Body                       = 
    
"Dear colleague...


"
  
}


## Send the message
Send-MailMessage @mailSpec

