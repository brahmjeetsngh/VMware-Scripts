#ReadMe:Connect to vCenter using Power-CLI and run this script

$cred = Get-Credential -UserName "root" -message "Enter new ESXi root password"
$vmhosts = get-vmhost | Out-GridView -PassThru -Title "Select ESXi hosts for changing the root password"
Foreach ($vmhost in $vmhosts) {
    $esxcli = get-esxcli -vmhost $vmhost -v2 
    $esxcli.system.account.set.Invoke(@{id=$cred.UserName;password=$cred.GetNetworkCredential().Password;passwordconfirmation=$cred.GetNetworkCredential().Password})
}