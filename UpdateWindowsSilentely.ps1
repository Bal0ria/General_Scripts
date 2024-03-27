#Powershell script to check and Install Windows update silently
#Process runs silentely without User intervention 
#Generates Log file
#Reboot is not enforced
#Build for DWS lab Use 

$Results = @(
@{ ResultCode = '0'; Meaning = "Not Started"}
@{ ResultCode = '1'; Meaning = "In Progress"}
@{ ResultCode = '2'; Meaning = "Succeeded"}
@{ ResultCode = '3'; Meaning = "Succeeded With Errors"}
@{ ResultCode = '4'; Meaning = "Failed"}
@{ ResultCode = '5'; Meaning = "Aborted"}
)
$InstallUpdates = "TRUE"
$WUDownloader=(New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
$WUInstaller=(New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
$WUUpdates=New-Object -ComObject Microsoft.Update.UpdateColl
$Global:RootDir = "C:\Windows\Temp"
$Global:LogDir = "$Global:RootDir\BuildLogs"
$Global:LogFile = "$Global:LogDir\WindowsUpdate.log"

Function _WriteLog ($Msg, $Type) {
    IF ($Type -eq "INFO") {Write-Host "$(Get-Date -format g)::INFO    :: $Msg" -ForegroundColor White; Out-File -FilePath "$Global:LogFile" -Append -InputObject "$(Get-Date -format g)::INFO    :: $Msg";}
    IF ($Type -eq "SUCCESS") {Write-Host "$(Get-Date -format g)::SUCCESS :: $Msg" -ForegroundColor Green; Out-File -FilePath "$Global:LogFile" -Append -InputObject "$(Get-Date -format g)::SUCCESS :: $Msg";}
    IF ($Type -eq "ERROR") {Write-Host "$(Get-Date -format g)::ERROR   :: $Msg" -ForegroundColor Red; Out-File -FilePath "$Global:LogFile" -Append -InputObject "$(Get-Date -format g)::ERROR   :: $Msg";}
    IF ($Type -eq "WARNING") {Write-Host "$(Get-Date -format g)::WARNING :: $Msg" -ForegroundColor Yellow; Out-File -FilePath "$Global:LogFile" -Append -InputObject "$(Get-Date -format g)::WARNING :: $Msg";}
}

((New-Object -ComObject Microsoft.Update.Session).CreateupdateSearcher().Search("IsInstalled=0 and Type='Software'")).Updates|%{
    if(!$_.EulaAccepted){$_.EulaAccepted=$true}
    if ($_.Title -notmatch "Preview"){[void]$WUUpdates.Add($_)}
}
 
if ($WUUpdates.Count -ge 1){
    if ($InstallUpdates -eq "TRUE"){
        $WUInstaller.ForceQuiet=$true
        $WUInstaller.Updates=$WUUpdates
        $WUDownloader.Updates=$WUUpdates
        $UpdateCount = $WUDownloader.Updates.count
        _WriteLog -Msg "Downloading $UpdateCount Updates" -Type "INFO";
        
            foreach ($update in $WUInstaller.Updates){_WriteLog -Msg "$($update.Title)" -Type "INFO"}
            $Download = $WUDownloader.Download()
            if ($Download.HResult -ne 0){
                $Convert = $Install.HResult
                $Hex = [System.Convert]::ToString($Convert, 16)
                $Hex = $Hex.Replace("ffffffff","0x")
                _WriteLog -Msg "Download HResult HEX: $Hex" -Type "INFO";
            }
        $InstallUpdateCount = $WUInstaller.Updates.count
        _WriteLog -Msg "Installing $InstallUpdateCount Updates" -Type "INFO";
        $Install = $WUInstaller.Install()
        $ResultMeaning = ($Results | Where-Object {$_.ResultCode -eq $Install.ResultCode}).Meaning
        _WriteLog -Msg "Result: $ResultMeaning" -Type "INFO";
        if ($Install.HResult -ne 0){
            $Convert = $Install.HResult
            $Hex = [System.Convert]::ToString($Convert, 16)
            $Hex = $Hex.Replace("ffffffff","0x")
            _WriteLog -Msg "Install HResult HEX: $Hex" -Type "INFO";
        }
        if ($Install.RebootRequired -eq $true){
            _WriteLog -Msg "Updates Require Restart" -Type "INFO";
        }
    }
    else
        {
        _WriteLog -Msg "Available Updates:" -Type "INFO";
        foreach ($update in $WUUpdates){_WriteLog -Msg  "$($update.Title)" -Type "INFO"}
     }
} 
else {_WriteLog -Msg "No updates detected" -Type "INFO";}
