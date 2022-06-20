Import-Module MicrosoftTeams
$userCredential = Get-Credential
Connect-MicrosoftTeams -Credential $userCredential

New-CsApplicationAccessPolicy -Identity MeetingAccessPolicy -AppIds "08be4766-f2bf-4af6-9157-19a7257483e7" -Description "Daemon can work with this user ids"

Get-CsOnlineUser | Grant-CsApplicationAccessPolicy -PolicyName MeetingAccessPolicy

#Remove-CsApplicationAccessPolicy -Identity MeetingAccessPolicy