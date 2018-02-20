<#
.Synopsis
Decrypt-IOS7

Author: Andy @ Netscylla (c) 2018
License: GNU GPL v3
Version: v0.1

.Description
Simple script decode/decrypt Cisco IOS Type 7 hashes

.Usage
.\Decrypt-IOS7.ps1 <hash>

.Example
.\Decrypt-IOS7.ps1 04480E051A33490E
#>

[CmdletBinding()]            
param( 
   [Parameter(Position = 0, Mandatory = $True)] [string] $hash           
)  

function Decrypt-IOS7{

$V = @(0x64, 0x73, 0x66, 0x64, 0x3b, 0x6b, 0x66, 0x6f, 0x41, 0x2c, 0x2e,
    0x69, 0x79, 0x65, 0x77, 0x72, 0x6b, 0x6c, 0x64, 0x4a, 0x4b, 0x44,
    0x48, 0x53, 0x55, 0x42, 0x73, 0x67, 0x76, 0x63, 0x61, 0x36, 0x39,
    0x38, 0x33, 0x34, 0x6e, 0x63, 0x78, 0x76, 0x39, 0x38, 0x37, 0x33,
    0x32, 0x35, 0x34, 0x6b, 0x3b, 0x66, 0x67, 0x38, 0x37);

  $pw=$hash                             
  $i=[int]::Parse($pw.substring(0,2));       # Initial index into Vigenere translation table
  $c=2;                                      # Initial pointer
  $r="";                                     # Variable to hold cleartext password
  while ($c -lt $pw.length){                 # Process each pair of hex values
    $x=([Convert]::toint16($pw.SubString($c, 2), 16))
    $y = $V[$i];
    $z=$x -bxor $y;                          # Vigenere reverse translation
    $r=$r+[char]$z
    $c+=2;    
    $i++;                                    # Move pointer to next hex pair
    $i%=53;                                  # Vigenere table wrap around
  }                                          
  Write-Host $r;    
}
Decrypt-IOS7;
