param(
    $expiresInDays=''
)

#$ErrorActionPreference = "stop"

$UserName = "xxx@xxx"
$AADSecurePassword = "XXX" | ConvertTo-SecureString -AsPlainText -Force
$AADCredential = New-Object PSCredential $UserName,$AADSecurePassword
Connect-AzureAD -Credential $AADCredential -TenantId "XXXX-XXXX-XXXX-XXXX-XXXXX"

$expirationDate = (Get-Date).AddDays($expiresInDays)
$currentDate = Get-Date
$AzureADApps = Get-AzureADApplication -All:$true
$AzureADApps = @() + $AzureADApps

$expiringSecrets = @()

foreach($app in $AzureADApps){
    $appCredentialsArr = @()

    $appCredentialsArr = Get-AzureADApplicationPasswordCredential -objectId $app.ObjectID | where-object {$_.EndDate -lt $expirationDate -and $_.EndDate -gt $currentDate}

    #$appCredentialsArr = Get-AzureADApplicationKeyCredential -objectId $app.ObjectID | where-object {$_.EndDate -lt $expirationDate -and $_.EndDate -gt $currentDate}                        

    foreach($credential in $appCredentialsArr){

        $credential | Select * | clip
        $expiringApp = [PSCustomObject] @{
                            AppName = $app.DisplayName
                            AppObjectID = $app.ObjectID
                            AppId = $app.AppId
                            KeyID = $credential.KeyId
                            ExpirationDate = $credential.EndDate
                        }
        
        $expiringSecrets += $expiringApp
    }
}

if($expiringSecrets.count -gt 0){
    $style = "<style>"
    $style = $style + "TABLE{font-size: 12px; border: 1px solid #CCC; font-family: Arial, Helvetica, sans-serif;}"
    $style = $style + "TH{margin: 4px; padding: 4px; background-color: #104E88; color: #FFF; font-weight: bold;}"
    $style = $style + "TD{padding: 4px; margin: 4px; border: 1px solid #CCC;}"
    $style = $style + "</style>"

    $reportHTML = $expiringSecrets | convertTo-Html -Head $style
    $reportHTML = "The following Azure AD Service Principal secrets will be expire in the next $expiresInDays days`n`n" + $reportHTML
    $reportHTML
}

$mailParams = @{
    smtpServer = 'smtp.office365.com'
    port = '587'
    UseSSL = $true
    credential = $AADCredential
    From = "xxx@xxx"
    To = "to@emailaddress"
    subject = "Expiring Azure AD Service Principal Secrets in $expiresInDays days"
    Body = $reportHTML
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

Send-MailMessage @mailParams -BodyAsHtml
