$AppId = ""
$TenantId = ""
 
Connect-AzureAd -Tenant $TenantId      
($Principal = Get-AzureADServicePrincipal -filter "AppId eq '$AppId'")
$PrincipalId = $Principal.ObjectId 
 
$DisplayName = "IMAP/smtp" 
 
Connect-ExchangeOnline
 
New-ServicePrincipal -AppId $AppId -ServiceId $PrincipalId -DisplayName $DisplayName
 
 Add-MailboxPermission -User $PrincipalId -AccessRights FullAccess -Identity "sharemialbox@contoso.com"
 
 
Add-RecipientPermission -Trustee $PrincipalId -AccessRights "SendAs" -Identity "sharemialbox@contoso.com"

