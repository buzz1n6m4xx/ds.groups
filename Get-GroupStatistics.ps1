$Path = "$env:USERPROFILE\Desktop\"

<# This PowerShell script either counts members of Active Directory based Dynamic Distribution Groups or exports members of Dynamic Distribution-, Distribution- or Security Groups #>


__author__ = "buzz1n6m4xx"
__status__ = "Production"
__version__ = "0.1"

# Define Menu Selection
function Show-Menu
{
    param (
        [string]$Title = 'Distribution-/Security Group Statistics'
    )
    Clear-Host
    Write-Host ("`r`n" + "         .oOo.   $Title   .oOo." + "`r`n") -ForegroundColor Cyan
    
    Write-Host "1: Press '1' to get member counts of all Dynamic DLs // Output on Screen" -ForegroundColor White
    Write-Host "2: Press '2' to get member counts of all Dynamic DLs // Output to Export File" -ForegroundColor White
    Write-Host "3: Press '3' to export members of specific Dynamic DLs (Input File" -ForegroundColor White -NoNewline
    Write-Host " DDL_Names.csv " -ForegroundColor Green -NoNewline
    Write-Host "must be located on the Desktop!)"  -ForegroundColor White
    Write-Host "4: Press '4' to export members of specific DLs (Input File" -ForegroundColor White -NoNewline
    Write-Host " DL_Names.csv " -ForegroundColor Green -NoNewline
    Write-Host "must be located on the Desktop!)" -ForegroundColor White
    Write-Host "5: Press '5' to export members of specific SGs (Input File" -ForegroundColor White -NoNewline
    Write-Host " SG_Names.csv " -ForegroundColor Green -NoNewline
    Write-Host "must be located on the Desktop!)" -ForegroundColor White
    Write-Host ("Q: Press 'Q' to quit." + "`r`n") -ForegroundColor White
    Write-Host "   Please note that" -ForegroundColor White -NoNewline
    Write-Host " Input Files " -ForegroundColor Green -NoNewline
    Write-Host "must have" -ForegroundColor White -NoNewline
    Write-Host " Name " -ForegroundColor Red -NoNewline
    Write-Host ("as column identifier/address!" + "`r`n") -ForegroundColor White

}

function Remove-Files
{
# Removing old Export Files in Path
Remove-Item -Path "$Path/DDL_Names_All.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "$Path/DDL_Count_All.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "$Path/DDL_Members.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "$Path/DL_Members.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "$Path/SG_Members.csv" -ErrorAction SilentlyContinue
}

Show-Menu –Title 'Distribution-/Security Group Statistics'
 $selection = Read-Host "Choose an option"
 switch ($selection)
 {
     '1' {

Remove-Files

# Getting ALL Dynamic Distribution Groups including Recipient Filter and Scope
Write-Host ("`r`n" + 'Generating Dynamic Distribution List Details...' + "`r`n") -ForegroundColor White -NoNewline
Get-DynamicDistributionGroup | Select-Object -Property Name, RecipientContainer, LdapRecipientFilter, RecipientFilterType | Export-Csv -Path "$Path/DDL_Names_All.csv" -NoTypeInformation

# Importing Input File
$Groups = Import-Csv -Path "$Path/DDL_Names_All.csv"

# Looping through each Group
Write-Host ("`r`n" + 'Calculating Number of Dynamic Distribution List Members...' + "`r`n") -ForegroundColor White -NoNewline 
ForEach ($Line in $Groups) 
{ 
$GroupName = $Line.Name
$GetGroup = Get-DynamicDistributionGroup -Identity $Line.Name

# Calculating Count of Members
$GetCount = (Get-Recipient -Resultsize Unlimited -OrganizationalUnit $GetGroup.RecipientContainer -RecipientPreviewFilter $GetGroup.RecipientFilter).count

# Writing Output to Host
Write-Host ("Groupname: " + $GroupName + "`n" + "Members  : " + $GetCount + "`r`n")
}

Remove-Item -Path "$Path/DDL_Names_All.csv" -ErrorAction SilentlyContinue

     } '2' {

Remove-Files

# Getting ALL Dynamic Distribution Groups including Recipient Filter and Scope

Write-Host ("`r`n" + 'Generating Dynamic Distribution List Details...' + "`r`n") -ForegroundColor White -NoNewline
Write-Host ('DDL Details have been exported to       : ') -ForegroundColor White -NoNewline
Write-Host ($Path + 'DDL_Names_All.csv') -ForegroundColor Green
Get-DynamicDistributionGroup | Select-Object -Property Name, RecipientContainer, LdapRecipientFilter, RecipientFilterType | Export-Csv -Path "$Path/DDL_Names_All.csv" -NoTypeInformation

# Importing Input File
$Groups = Import-Csv -Path "$Path/DDL_Names_All.csv"

# Looping through each Group
Write-Host ("`r`n" + 'Calculating Number of Dynamic Distribution List Members...' + "`r`n") -ForegroundColor White -NoNewline 
ForEach ($Line in $Groups) 
{ 
$GroupName = $Line.Name
$GetGroup = Get-DynamicDistributionGroup -Identity $Line.Name

# Calculating Count of Members
$GetCount = (Get-Recipient -Resultsize Unlimited -OrganizationalUnit $GetGroup.RecipientContainer -RecipientPreviewFilter $GetGroup.RecipientFilter).count

# Saving Output for Member Count
$GetStats = @()
$GetStats += New-Object -TypeName PSObject -Property @{GroupName=$GroupName;Members=$GetCount}

# Writing Member Count Output to CSV File
$GetStats | Export-Csv -Path "$Path/DDL_Count_All.csv" -Delimiter ';' -Append -NoTypeInformation

}
Write-Host ('DDL Member Counts have been exported to : ') -ForegroundColor White -NoNewline
Write-Host ($Path + 'DDL_Count_All.csv' + "`r`n`r`n") -ForegroundColor Green -NoNewline

     } '3' {

Remove-Files

# Importing Input File
Write-Host ("`r`n" + 'Importing Dynamic Distribution List Details...' + "`r`n") -ForegroundColor White -NoNewline
$Groups = Import-Csv -Path "$Path/DDL_Names.csv"

# Looping through each Group
Write-Host ("`r`n" + 'Enumerating Dynamic Distribution List Members...' + "`r`n") -ForegroundColor White -NoNewline 
ForEach ($Line in $Groups) 
{ 
$GroupName = $Line.Name
$GetGroup = Get-DynamicDistributionGroup -Identity $Line.Name

#Enumerating and Exporting Group Members
$GetMembers = Get-Recipient -Resultsize Unlimited -OrganizationalUnit $GetGroup.RecipientContainer -RecipientPreviewFilter $GetGroup.RecipientFilter | Select {write-output "$GroupName"}, PrimarySmtpAddress | Export-Csv -Path "$Path/DDL_Members.csv" -NoTypeInformation -Append

}
Write-Host ('DDL Members have been exported to : ') -ForegroundColor White -NoNewline
Write-Host ($Path + 'DDL_Members.csv' + "`r`n`r`n") -ForegroundColor Green -NoNewline

     } '4' {

Remove-Files

# Importing Input File
Write-Host ("`r`n" + 'Importing Distribution List Details...' + "`r`n") -ForegroundColor White -NoNewline
$Groups = Import-Csv -Path "$Path/DL_Names.csv"

# Looping through each Group
Write-Host ("`r`n" + 'Enumerating Distribution List Members...' + "`r`n") -ForegroundColor White -NoNewline 
ForEach ($Line in $Groups) 
{ 
$GroupName = $Line.Name

#Enumerating and Exporting Group Members
$GetMembers = Get-DistributionGroupMember -Identity $Line.Name | Select {write-output "$GroupName"}, PrimarySmtpAddress | Export-Csv -Path "C:\Users\a-hoerthm\Desktop\DL_Members.csv" -NoTypeInformation -Append

}
Write-Host ('DL Members have been exported to : ') -ForegroundColor White -NoNewline
Write-Host ($Path + 'DL_Members.csv' + "`r`n`r`n") -ForegroundColor Green -NoNewline

     } '5' {

Remove-Files

# Importing Input File
Write-Host ("`r`n" + 'Importing Security Group Details...' + "`r`n") -ForegroundColor White -NoNewline
$Groups = Import-Csv -Path "$Path/SG_Names.csv"

# Looping through each Group
Write-Host ("`r`n" + 'Enumerating Security Group Members...' + "`r`n") -ForegroundColor White -NoNewline 
ForEach ($Line in $Groups) 
{ 
$GroupName = $Line.Name

#Enumerating and Exporting Group Members
$GetMembers = Get-ADGroupMember -Identity $Line.Name | Select {write-output "$GroupName"}, displayName, Name, sAMAccountName | Export-Csv -Path "C:\Users\a-hoerthm\Desktop\SG_Members.csv" -NoTypeInformation -Append

}
Write-Host ('SG Members have been exported to : ') -ForegroundColor White -NoNewline
Write-Host ($Path + 'SG_Members.csv' + "`r`n`r`n") -ForegroundColor Green -NoNewline

     } 'q' {
        return
     }
}