[OutputType("PSAzureOperationResponse")]
param
(
    [Parameter (Mandatory=$false)]
    [object] $WebhookData
)
$ErrorActionPreference = "stop"
if ($WebhookData)
{
    # Get the data object from WebhookData
    $WebhookBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    #$WebhookBody = $WebhookData | ConvertFrom-Json
    $Json = ($WebhookBody.data.alertContext.SearchResults.tables.rows).trim("{}")
}
foreach ($J in $Json){
$Properties = @()
$UserCredential = Get-AutomationPSCredential -Name 'Email'
# Application (client) ID, tenant Name and secret
$clientId = Get-AutomationVariable -Name 'ClientID'
$tenantName = Get-AutomationVariable -Name 'TenantID'
$clientSecret = Get-AutomationVariable -Name 'Secret'
$EmailAddress = Get-AutomationVariable -Name 'Address'
$resource = "https://graph.microsoft.com/"
#Get Token
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

#Get all succesfully created users for the last 1hr
$GetDisplayName = "https://graph.microsoft.com/beta/accessReviews/$($J)"
$o = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $GetDisplayName -Method Get 
$GetDecisions = "https://graph.microsoft.com/beta/accessReviews/$($J)/decisions"
$g = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $GetDecisions -Method Get 
for ($i = 0 ; $i -lt $g.value.count; $i++){
if ($g.value[$i].reviewResult -eq "NotReviewed"){
$Properties += [PSCustomObject]@{
"Display Name" = $g.value[$i].userDisplayName
UserprincipalName = $g.value[$i].userPrincipalName
"Reviewed Group" = $o.reviewedEntity.displayName
Result = $g.value[$i].reviewResult

}
}
$Properties
}
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
$Header = Out-String -InputObject ($Properties | ConvertTo-Html -head $style -Body "<font color=`"Black`"><h4><left>Not Reviewed</left></h4></font>")
$mailParams = @{
    SmtpServer                 = 'smtp.office365.com'
    Port                       = '587' # or '25' if not using TLS
    UseSSL                     = $true ## or not if using non-TLS
    Credential                 = $UserCredential
    From                       = $UserCredential.UserName
    To                         = $EmailAddress
    Subject                    = "Access Review"
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}
if ($Properties -ne $null) { Send-MailMessage @mailParams -Body $Header -BodyAsHtml } 
}

