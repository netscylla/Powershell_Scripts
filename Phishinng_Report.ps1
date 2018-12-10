# Phishing Report v0.1
# Copyright Netscylla 2018
# Version 0.1 10/12/2018
# Used to collect email artifacts in specific OWA folder, currently: Email address, Name 

$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");

#Find inbox with case/campaign
#example = "\\Phishing\Phishing Campaign 1"
#use following command to locate mailbox/folder: $ns.Folders[4].Folders | select folderpath;
$inbox = $ns.Folders.Item(3).Folders.Item(6)

$inbox.items | foreach {

    # convert Exchange user account into human readable smtp email address
    if ($_.SenderEmailType.ToUpper().Equals("EX")){
        $recip = $ns.CreateRecipient($_.SenderName);
        $exUser = $recip.AddressEntry.GetExchangeUser();
        $smtpAddress = $exUser.PrimarySmtpAddress;
    }else{
        $smtpAddress = mi.SenderEmailAddress;
    }
    # you may want to comment this out depending on reporting requirements?
    "$smtpAddress,$($_.SenderName)" | out-file c:\Temp\Report.csv -Append
}
#uncomment if you want Powershell to dump a table
#$inbox.items| Select SenderEmailAddress,SenderName |Format-Table -AutoSize
