<#
.Synopsis
Phishing Report v0.2

Author: Andy @ Netscylla (c) 2018
License: GNU GPL v3
Version: v0.2

.Description
Simple script to collect email artifacts in specific OWA folder, currently: Email address, Name 
#>

param(  
[Parameter(Position = 0, Mandatory = $false)] [string] $filename = "c:\Report.csv",      
[Parameter(Position = 0, Mandatory = $false)] [string] $file = $false         
)   
 
#$search_folder = "\\Phishing\Phishing Campaign"
$search_folder = Read-Host -Prompt 'Enter mail folder (ex: \\Phishing\Phishing Campaign)'
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
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
 
if (!$inbox){
  exit
}
 
if ($file -eq $false){
  Write-Host "Dumping to console"
  Write-Host "Email,Surname,FirstName"
}else{
  Write-Host "Dumping to file " $filename
}
 
$inbox.items | foreach {
    if ($_.SenderEmailType.ToUpper().Equals("EX")){
      $recip = $ns.CreateRecipient($_.SenderName);
      if ($recip.AddressEntry){      
          $exUser = $recip.AddressEntry.GetExchangeUser();
          $smtpAddress = $exUser.PrimarySmtpAddress;
      }else{
          $namearray=@()
          $namearray=$_.SenderName.Split(" ,",[System.StringSplitOptions]::RemoveEmptyEntries)
          [array]::Reverse($namearray)
          $email_prefix= $namearray -Join "."
          $smtpAddress=$email_prefix + "@domain.com"
      }
    }else{
      $smtpAddress = $_.SenderEmailAddress;
    }
 
    if($file -eq 1) {
      "$smtpAddress,$($_.SenderName)" | out-file $filename -Append 
    }else{
      "$smtpAddress,$($_.SenderName)" |Format-Table -Autosize
    }
    $smtpAddress=""
}
