# connect to Microsoft Exchange

Set-ExecutionPolicy RemoteSigned -Force

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session

# view calendar permissions
# Get-MailboxFolderPermission -Identity <user's email whose calendar permissions you want to view>:\Calendar

Add-MailboxFolderPermission -Identity <email of the object of the permissions>:\calendar -user <email of person getting permissions> -AccessRights Reviewer

