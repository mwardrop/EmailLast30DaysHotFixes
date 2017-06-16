Function EmailLast30HotFix
{

    $outputs = Invoke-Expression "wmic qfe list" 
    $outputs = $outputs[1..($outputs.length)] 
     
    $date1 = Get-Date -Date "01/01/1970"
    
    $list = @()
     
    foreach ($output in $Outputs) { 
        if ($output) { 
            $output = $output -replace 'y U','y-U' 
            $output = $output -replace 'NT A','NT-A' 
            $output = $output -replace '\s+',' ' 
            $parts = $output -split ' ' 
            if($parts[5]) {
                if ($parts[5] -like "*/*/*") { 
                    $Dateis = [datetime]::ParseExact($parts[5], '%M/%d/yyyy',[Globalization.cultureinfo]::GetCultureInfo("en-US").DateTimeFormat) 
                } else { 
                    $Dateis = get-date([DateTime][Convert]::ToInt64("$parts[5]", 16))-Format '%M/%d/yyyy' 
                } 
                $list += New-Object -Type PSObject -Property @{ 
                    KBArticle = [string]$parts[0] 
                    Computername = [string]$parts[1] 
                    Description = [string]$parts[2] 
                    FixComments = [string]$parts[6] 
                    HotFixID = [string]$parts[3] 
                    InstalledOn = Get-Date($Dateis)-format "yyyy/MMMM" 
                    InstalledBy = [string]$parts[4] 
                    InstallDate = [string]$parts[7] 
                    Name = [string]$parts[8] 
                    ServicePackInEffect = [string]$parts[9] 
                    Status = [string]$parts[10] 
                    Epoch = [Decimal](New-TimeSpan -Start $date1 -End $Dateis).TotalSeconds
                } 
            }

        } 
    } 



    $list = $list | Where-Object {$_.Epoch -gt (New-TimeSpan -Start (Get-Date -Date "01/01/1970") -End (get-date).AddDays(-30)).TotalSeconds}| Sort-Object -Property Epoch -Descending | FT InstalledOn,HotFixID,ComputerName

    $list

}

Clear

$hosts = @("HOST1", "HOST2", "Host3")

$emailBody = @()

for ($i=0; $i -lt $hosts.length; $i++) { 

    $remote += Invoke-Command -ComputerName $hosts[$i] ${function:EmailLast30HotFix}|ft -HideTableHeaders  

    Write-Host "Hot Fixes Retrieved From:" $hosts[$i] -ForegroundColor green

    $remote | Out-String

    $emailBody += $remote

} 

$emailDate = Get-Date -format "yyyy/MMMM"

$emailSubject = "$emailDate Windows Updates last 30 days"

$email = @{
    From = "xxx@yyy.com"
    To = "xxx@yyy.com"
    Subject = $emailSubject
    SMTPServer = "smtp.yyy.com"
    Body = $emailBody | Out-String
}

send-mailmessage @email