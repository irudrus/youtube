
Set-ExecutionPolicy RemoteSigned
Connect-ExchangeOnline   

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$filePath = 'C:\Temp\SendAsPermissions24MarzFixed.csv'
$mailboxes = Get-Mailbox -ResultSize Unlimited
$sendAsPermissions = @()
foreach($mailbox in $mailboxes){
   $ADPermissions = $null
   $ADPermissions = Get-ADPermission $mailbox.DistinguishedName | ? { $_.ExtendedRights -like "*send*" -and -not ($_.User -match "NT AUTHORITY")}
   if ($ADPermission -ne $null){
       $userRecipient = $null
       foreach ($ADPermission in $ADPermissions){
       $userRecipient = Get-Recipient $ADPermission.User.RawIdentity.Split("\")[1]
       $sendAsPermissions += $ADPermission | ? { $_.ExtendedRights -like "*send*" -and -not ($_.User -match "NT AUTHORITY")} | Select-Object @{name='WindowsEmailAddress';expression={$mailbox.WindowsEmailAddress}},@{name='Identity';expression={$_.Identity}},@{name='User';expression={$_.User}},@{name='UserPrimarySmtpAddress';expression={$userRecipient.PrimarySmtpAddress.Address}},@{name='Deny';expression={$_.Deny}}
       } 
   }
}


$filePath = 'C:\temp\SendAsPermissions2.csv'
$mailboxesfilePath = 'C:\temp\December12.csv'
$sendAsPermissions = @()
$sendAsPermissions = Import-Csv -Path $filePath
$movedmailboxes = @()
$movedmailboxes = Import-Csv -Path $mailboxesfilePath

foreach($movedmailbox in $movedmailboxes){
    foreach($sendAsPermission in $sendAsPermissions){
        if($movedmailbox.EmailAddress -eq $sendAsPermission.UserPrimarySmtpAddress){
              
        Add-RecipientPermission $sendAsPermission.WindowsEmailAddress  -AccessRights SendAs -Trustee $movedmailbox.EmailAddress -Confirm:$False
   
         }
    }
  
}


