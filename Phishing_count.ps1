<#
.Synopsis
Phishing Count v0.1
Author: Andy @ Netscylla (c) 2019
License: GNU GPL v3
Version: v0.1

.Description
Simple script to count pre-defined categories in a mail fodler to collect statistics on phishing and spam emails reported via customers and/or internal staff 
* Green Category  = Genuine email
* Red Category    = Phishing email
* Purple Category = Malware inside email
* Blue Category   = Spam email 
Outlook can tag emails with colour categories, refer to this article for more information
* https://support.office.com/en-us/article/create-and-assign-color-categories-a1fde97e-15e1-4179-a1a0-8a91ef89b8dc
#>

#Our phishing inbox and triage folder
$search_folder = "\\Phishing\triage"
#Find the triage folder location in powershell
$outlook = new-object -com outlook.application
$month = Read-Host -Prompt 'Enter month as integer (e.g. 12 = dec)'

$ns = $outlook.GetNameSpace("MAPI")
$j=0
for ($i=1; $i -lt 4; $i++){
  $ns.Folders[$i].Folders|ForEach-Object{
    $j+=1
    if ($_.FolderPath -eq $search_folder){
      $inbox = $ns.Folders.Item($i).Folders.Item($j)
      Write-Host -ForegroundColor Green "Found"
      break
    }else{
      Write-host -NoNewLine "."
    }
  }
  $j=0
}
 
#if folder empty/doesnt exist exit
if (!$inbox){
  exit
}
#clear variables
$blue_count=0;
$red_count=0;
$green_count=0;
$purple_count=0;

#Process emails
$inbox.items | foreach {
    $dt=[DateTime]($_.ReceivedTime).datetime
    if ($dt.Month -eq $month){
      if ($_.Categories -eq "Blue Category"){
        $blue_count++
      }
      elseif ($_.Categories -eq "Red Category"){
        $red_count++
      }
      elseif ($_.Categories -eq "Green Category"){
        $green_count++
      }
      elseif ($_.Categories -eq "Purple Category"){
        $purple_count++
      }
      else{
        #Once we've finished for the queried month, stop reading additional mail.
        if ( ($dt.Month%12) -le ($month - 1)%12){return}
      }
    }
}

$total=$blue_count + $red_count +  $purple_count + $green_count
$bp=[math]::Round($blue_count/$total*100,2)
$rp=[math]::Round($red_count/$total*100,2)
$pp=[math]::Round($purple_count/$total*100,2)
$gp=[math]::Round($green_count/$total*100,2)

write-host "spam count   : $blue_count"
write-host -ForegroundColor Red "phish count  : $red_count"
write-host "malware count: $purple_count"
write-host -ForegroundColor Green "genuine count: $green_count"
write-host "Total count  : $total"

write-host "spam %       : $bp %"
write-host -ForegroundColor Red "phish %      : $rp %"
write-host "malware %    : $pp %"
write-host -ForegroundColor Green "genuine %    : $gp %"
