
<#
.Synopsis
Phishing Move v0.1
Author: Andy @ Netscylla (c) 2019
License: GNU GPL v3
Version: v0.1
.Description
Simple script to search mail for a specific Subject or attachment filename and move to a folder for appropriate processing/triage 
#>

# Insert your Inbox that receives phishing email here: 
$search_folder = "\\Phishing\Inbox"
# Insert your triage folder here:
$target_folder = "\\Phishing\xxx"
# Insert the string (Subject line) that you wish to match
$bad_string = "xxx"

# setup mapi interface 
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");

$j=0
$found_boxes=0;
for ($i=1; $i -lt 5; $i++){
  $ns.Folders[$i].Folders|ForEach-Object{
    $j+=1
    if ($_.FolderPath -eq $search_folder){
      $inbox = $ns.Folders.Item($i).Folders.Item($j)
      Write-Host -ForegroundColor Green "Found inbox"
      Write-Host ""
      $found_boxes++
      if ($found_boxes -eq 2) { break }
    }else{
      Write-host -NoNewLine "."
    }
    if ($_.FolderPath -eq $target_folder){
      $mvbox = $ns.Folders.Item($i).Folders.Item($j)
      Write-Host -ForegroundColor Green "Found mv_box"
      Write-Host ""
      $found_boxes++
      if ($found_boxes -eq 2) { break }
    }else{
      #Write-host -NoNewLine "."
    }
  }
  $j=0
}

Write-Host ""

#If inbox was not found, exit program 
if (!$inbox){
  Write-host -ForegroundColor Red "Error: Inbox/Folder Not Found!"
  exit

}
#If inbox was not found, exit program 
if (!$mvbox){
  Write-host -ForegroundColor Red "Error: Destination Folder Not Found!"
  exit

}
 
# Search unread mail for our search criteria '$bad_string' if we find a match, move the email to the destination/triage folder
$restricteditems=$inbox.items.Restrict("[Unread] = true")
for ($inc=$restricteditems.count; $inc -gt 0 ; $inc--){
  if ($restricteditems.item($inc).Subject -match $bad_string){
    try{
      write-host "Moving mail to IPR, mail from:", $restricteditems.item($inc).SenderName
      [void]$restricteditems.item($inc).Move($mvbox)
    }catch{
      write-host "failed to move mail from:", $restricteditems.item($inc).SenderName
    }
  }
  if(($restricteditems.item($inc).Attachments|select FileName) -match $bad_string){
    try{
      write-host "Moving mail to IPR, mail from:", $restricteditems.item($inc).SenderName
      [void]$restricteditems.item($inc).Move($mvbox)
    }catch{
      write-host "failed to move mail from:", $restricteditems.item($inc).SenderName
    }
  }
}
