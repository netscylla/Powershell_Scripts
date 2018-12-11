<#
.Synopsis
Phishing Report v0.1

Author: Andy @ Netscylla (c) 2018
License: GNU GPL v3
Version: v0.1

.Description
Simple script to collect email artifacts in specific OWA folder, currently: Email address, Name 
#>

$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");

#Find inbox with case/campaign
#example = "\\Phishing\Phishing Campaign 1"
#use following command to locate mailbox/folder: $ns.Folders[4].Folders | select folderpath;
$inbox = $ns.Folders.Item(3).Folders.Item(6)

$inbox.items | foreach {
    if ($_.SenderEmailType.ToUpper().Equals("EX")){
      $recip = $ns.CreateRecipient($_.SenderName);
      #if address entry exists continue
      if ($recip.AddressEntry){     
          $exUser = $recip.AddressEntry.GetExchangeUser();
          $smtpAddress = $exUser.PrimarySmtpAddress;
      }else{
          #if address entry == null, guess at email address from name firstname.surname
          $namearray=@()
          $namearray=$_.SenderName.Split(" ,",[System.StringSplitOptions]::RemoveEmptyEntries)
          [array]::Reverse($namearray)
          $email_prefix= $namearray -Join "."
          $smtpAddress=$email_prefix + "@domain.com"
      }
    }else{
      $smtpAddress = $_.SenderEmailAddress;
    }
    "$smtpAddress,$($_.SenderName)" | out-file h:\test2.csv -Append
    $smtpAddress=""
}

#$inbox.items| Select SenderEmailAddress,SenderName |Format-Table -AutoSize
