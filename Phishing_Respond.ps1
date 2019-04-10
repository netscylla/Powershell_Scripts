<#
.Synopsis
Phishing Respond v0.1
Author: Andy @ Netscylla (c) 2019
License: GNU GPL v3
Version: v0.1
.Description
Simple script to search mail for a specific Subject, move to a folder for appropriate processing/triage, and track the reporter(s). 
Then send an email reply to all reporters (bcc), acknowledging the phishing campaign, or a message of your choosing. 
#>

$company_domain = "example.com"
$search_folder = "\\Phishing\Inbox"
$target_folder = "\\Phishing\Triage"
$bad_string = "Invoice Payment 7053" 

# function to bcc a string of email addresses, with a static message
function send_mail([string]$senderslist){
  try{
    write-host "Sending response mail to ", $senderslist
    $Mail = $outlook.CreateItem(0)
    $Mail.bcc = $replymail
    $Mail.Subject = "RE: " + $badstring + "Phishing Campaign"
    $Mail.Body = "Thank you for your concern,`nThe email was part of a phishing awareness campaign.`nOnce reported to Phishing@$company_domain feel free to ignore the email.`nThank you`nThis is an automated response!"
    $Mail.Send()

  }catch{
    write-host -Foreground Red "Sending response mail failed: ", $senderslist
  }
}
 

# configure MAPI object, and locate mail boxes
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
      $found++
      if ($found -eq 2 ) { break }  
    }else{
      if ($found -eq 2 ) { break }
      Write-host -NoNewLine "."
    }

    if ($_.FolderPath -eq $target_folder){
      $mvbox = $ns.Folders.Item($i).Folders.Item($j)
      Write-Host -ForegroundColor Green "Found mv_box"
      Write-Host ""
      $found++
      if ($found -eq 2 ) { break }
    }
  }
  $j=0
}
Write-Host ""

#if mailbox not found quit
if (!$inbox){
  exit
}

# perform our mailbox analysis below....
$senderslist="";
$restricteditems=$inbox.items.Restrict("[Unread] = true")
 

for ($inc=$restricteditems.count; $inc -gt 0 ; $inc--){
  #match subject name
  if ($restricteditems.item($inc).Subject -match $bad_string){
    try{
      $replymail= (($ns.CreateRecipient( $restricteditems.item($inc).SenderEmailAddress )).AddressEntry.GetExchangeUser()).PrimarySmtpAddress; 
      write-host "Moving mail to IPR, mail from:", $restricteditems.item($inc).SenderName
      [void]$restricteditems.item($inc).Move($mvbox)
    }catch{
      write-host -Foreground Red "failed to move mail from:", $restricteditems.item($inc).SenderName
    }
  }

  #match subject name as mail attachment
  try{ 
    if(($restricteditems.item($inc).Attachments|select FileName) -match $bad_string){
      try{
        $replymail= (($ns.CreateRecipient( $restricteditems.item($inc).SenderEmailAddress )).AddressEntry.GetExchangeUser()).PrimarySmtpAddress; 
        write-host "Moving mail to IPR, mail from:", $restricteditems.item($inc).SenderName
        [void]$restricteditems.item($inc).Move($mvbox)
      }catch{
        write-host -Foreground Red "failed to move mail from:", $restricteditems.item($inc).SenderName
      }
    }
  }catch {}

  #remove COM objects in reply string, and concatenate all email addresses tracked
  if ($replaymail -notcontains "Microsoft.Office"){
    $senderslist += $replymail + ";"
  }
}

send_mail($senderslist)
