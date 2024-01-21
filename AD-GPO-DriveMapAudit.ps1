### Modify These Variables ###
# Specify the OU
$ou = "OU=Security Groups,DC=domain,DC=local" # Ex: "OU=Security Groups,DC=domain,DC=local"
# Specify the name of the GPO
$gpoName = 'Shared Drives'
##############################

# Import Active Directory Module
Import-Module ActiveDirectory

# Get the domain name in uppercase and without the top-level domain
$domain = (Get-ADDomain).DNSRoot.Split('.')[0].ToUpper()

# Define the Domain Users group using the domain name
$domainUsers = "$domain\Domain Users"

# Get all the groups in the OU
$groups = Get-ADGroup -Filter * -SearchBase $ou

# Create an empty array to hold the results
$results = @()

foreach ($group in $groups) {
    # Get the members of the group
    $groupMembers = Get-ADGroupMember -Identity $group | ForEach-Object {
        # Check if the member is a user and is enabled
        if ($_.objectClass -eq 'user') {
            $user = Get-ADUser -Identity $_.distinguishedName -Properties Enabled
            if ($user.Enabled -eq $true) {
                $_.Name
            }
        } else {
            $_.Name
        }
    }

    # Sort the members alphabetically by first name
    $sortedMembers = $groupMembers | Sort-Object

    # Add the group name and its members to the results
    $results += New-Object PSObject -Property @{
        'GroupName' = "$domain\$($group.Name)"
        'Members' = $sortedMembers -join ', '
    }
}

# Import the Group Policy module
Import-Module GroupPolicy

# Get the GPO
$gpo = Get-GPO -Name $gpoName

# Export the GPO to an XML report
$xmlReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml

# Load the XML content
[xml]$xmlContent = $xmlReport

# Define the namespace manager
$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
$nsManager.AddNamespace("ns", "http://www.microsoft.com/GroupPolicy/Settings/DriveMaps")

# Define an empty array to hold the data
$csvData = @()

# Loop over each 'ns:Drive' element
foreach ($drive in $xmlContent.DocumentElement.SelectNodes('.//ns:Drive', $nsManager)) {
    $properties = $drive.SelectSingleNode("ns:Properties", $nsManager)
    $filter_groups = $drive.SelectNodes("ns:Filters/ns:FilterGroup", $nsManager)

    # Extract the required information
    $path = $properties.path
    $label = $properties.label
    $letter = $properties.letter
    $action = $properties.action

    # Map the action value to the corresponding text
    switch ($action) {
        'U' { $action = 'Update' }
        'C' { $action = 'Create' }
        'D' { $action = 'Delete' }
        'R' { $action = 'Replace' }
        default { $action = $action } # leave as is if none of the above
    }

    foreach ($filter_group in $filter_groups) {
        $groupName = $filter_group.name

        # Create a custom object for this row
        $row = New-Object PSObject -Property @{
            'Path' = $path
            'Label' = $label
            'Letter' = $letter
            'GroupName' = $groupName
            'Action' = $action # add the action
        }

        # Add the row to the CSV data
        $csvData += $row
    }

    if ($filter_groups.Count -eq 0) {
        # Create a custom object for this row with 'Everyone' as GroupName
        $row = New-Object PSObject -Property @{
            'Path' = $path
            'Label' = $label
            'Letter' = $letter
            'GroupName' = 'Everyone'
            'Action' = $action # add the action
        }

        # Add the row to the CSV data
        $csvData += $row
    }
}

# Merge the two data arrays
$mergedData = foreach ($csvDataRow in $csvData) {
    $matchingGroupMembersRow = $results | Where-Object { $_.GroupName -eq $csvDataRow.GroupName }
    if ($matchingGroupMembersRow) {
        Write-Host "Found matching GroupName: $($csvDataRow.GroupName)"

        # Create a new object that combines the data from the matching rows
        # Reorder the properties in the desired order
        $members = if ($matchingGroupMembersRow.GroupName -eq $domainUsers) {
            'Domain Users'
        } else {
            $matchingGroupMembersRow.Members
        }
        New-Object PSObject -Property @{
            'Letter' = $csvDataRow.Letter
            'Label' = $csvDataRow.Label
            'Path' = $csvDataRow.Path
            'GroupName' = $csvDataRow.GroupName
            'Members' = $members
            'Action' = $csvDataRow.Action # add the action
        }
    }
    else {
        Write-Host "No matching GroupName found for: $($csvDataRow.GroupName)"
        $members = if ($csvDataRow.GroupName -imatch 'Domain Admins') {
            $domainAdmins = Get-ADGroupMember -Identity 'Domain Admins' | ForEach-Object {
                if ($_.objectClass -eq 'user') {
                    $user = Get-ADUser -Identity $_.distinguishedName -Properties Enabled
                    if ($user.Enabled -eq $true) {
                        $_.Name
                    }
                } else {
                    $_.Name
                }
            }
            # Sort the domain admins members alphabetically by name
            $domainAdmins = $domainAdmins | Sort-Object
            $domainAdmins -join ', '
        } elseif ($csvDataRow.GroupName -eq $domainUsers) {
            'Domain Users'
        } else {
            'Everyone'
        }

        # Create a new object with the data from the current row, but with the members of 'Domain Admins' if applicable
        New-Object PSObject -Property @{
            'Letter' = $csvDataRow.Letter
            'Label' = $csvDataRow.Label
            'Path' = $csvDataRow.Path
            'GroupName' = $csvDataRow.GroupName
            'Members' = $members
            'Action' = $csvDataRow.Action # add the action
        }
    }
}

# Sort the merged data by the 'Letter' column in alphabetical order
$sortedData = $mergedData | Sort-Object Letter

# Define the sortedData and the fields to select
$selectedData = $sortedData | Select-Object Letter, Label, Action, Path, GroupName, Members

# Get current date and time
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Define the file paths
$xlsxFilePath = "C:\temp\"+$domain+"_MappedDrives_$timestamp.xlsx"
$csvFilePath = "C:\temp\"+$domain+"_MappedDrives_$timestamp.csv"

# Check if the "ImportExcel" module is installed
if (Get-Module -ListAvailable -Name ImportExcel) {
    # If the module is installed, export to an XLSX file
    $selectedData | Export-Excel -Path $xlsxFilePath -TableName 'MappedDrives' -TableStyle Medium6 -AutoSize
} else {
    try {
        # If the module is not installed, try to install it
        Install-Module -Name ImportExcel -Scope CurrentUser -Force

        # If the module installs successfully, export to an XLSX file
        $selectedData | Export-Excel -Path $xlsxFilePath -TableName 'MappedDrives' -TableStyle Medium6 -AutoSize
    } catch {
        # If the module fails to install, export to a CSV file
        $selectedData | Export-Csv -Path $csvFilePath -NoTypeInformation
    }
}