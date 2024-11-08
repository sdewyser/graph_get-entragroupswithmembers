# Graph_Get-EntraGroupsWithMembers.psm1

# Define the module manifest
function Get-ModuleManifest {
    @{
        RootModule    = 'Graph_Get-EntraGroupsWithMembers.psm1'
        ModuleVersion = '1.1.0'
        GUID          = 'b72ac9d4-8121-4f9d-b02e-31f4b9d1a2cd'
        Author        = 'Stefaan Dewulf'
        Description   = 'A PowerShell module to retrieve EntraID AD groups and their nested distinct members using Microsoft Graph API.'
        CompanyName   = 'dewyser.net'
    }
}
# Import necessary modules
Import-Module Microsoft.Graph

# Function to authenticate with Microsoft Graph using client credentials
function Connect-MicrosoftGraphAPI {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    # Create the authentication token
    $SecureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force

    $ClientCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecureClientSecret

    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientCredential -NoWelcome
    }
    catch {
        Write-Output "Error connecting to Microsoft Graph API."
    }
}

# Recursive function to retrieve all distinct users for a given group, including nested groups
function Get-AllDistinctGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Hashtable]$UserCache,  # Cache users by their ID to ensure uniqueness
        [bool]$IsTopLevelGroup  # Flag to indicate if this is a top-level group
    )

    $users = @()  # Array to hold distinct users for the group

    # If it's a top-level group, reset the cache
    if ($IsTopLevelGroup) {
        $UserCache.Clear()  # Reset the cache for a new top-level group
    }

    # Retrieve members with explicit properties (id, displayName, userPrincipalName, mail)
    $rawMembers = Get-MgGroupMember -GroupId $GroupId -All -Property 'id,displayName,userPrincipalName,mail'

    Write-Host "Processing Group: $GroupId - Total Members: $($rawMembers.Count)"

    foreach ($member in $rawMembers) {
        $memberType = $member.AdditionalProperties.'@odata.type'  # Automatically included by Graph
        $displayName = $member.DisplayName
        $userPrincipalName = $member.UserPrincipalName

        Write-Host "Processing Member: $($displayName) - Type: $memberType"

        if ($memberType -eq "#microsoft.graph.user") {
            # If missing displayName or userPrincipalName, query the user details explicitly
            if (-not $displayName -or -not $userPrincipalName) {
                Write-Host "Fetching additional details for user: $($member.Id)"
                $userDetails = Get-MgUser -UserId $member.Id -Property 'displayName,userPrincipalName,mail'
                $displayName = $userDetails.DisplayName
                $userPrincipalName = $userDetails.UserPrincipalName
            }

            if ($displayName -and $userPrincipalName) {
                # Ensure the user is only added once across all groups by using the UserCache
                if (-not $UserCache.ContainsKey($member.Id)) {
                    Write-Host "Adding User: $($displayName) ($($userPrincipalName))"
                    $users += [PSCustomObject]@{
                        Id               = $member.Id
                        DisplayName      = $displayName
                        UserPrincipalName = $userPrincipalName
                        Mail             = $userDetails.Mail
                    }
                    $UserCache[$member.Id] = [PSCustomObject]@{
                        Id               = $member.Id
                        DisplayName      = $displayName
                        UserPrincipalName = $userPrincipalName
                        Mail             = $userDetails.Mail
                    }  # Cache the full user object to avoid duplicates
                } else {
                    Write-Host "User already cached: $($displayName)"
                }
            } else {
                Write-Host "Skipped User due to missing properties: DisplayName or UserPrincipalName"
            }
        }
        elseif ($memberType -eq "#microsoft.graph.group") {
            # Process nested groups recursively, pass the current cache along
            Write-Host "Processing Nested Group: $($member.DisplayName) ($($member.Id))"
            $nestedUsers = Get-AllDistinctGroupMembers -GroupId $member.Id -UserCache $UserCache -IsTopLevelGroup $false
            $users += $nestedUsers
        } else {
            Write-Host "Skipped Member (Unknown Type): $($displayName)"
        }
    }

    return $users
}

# Function to retrieve all groups with their distinct members and count
function Get-AllGroupsWithDistinctMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string[]]$GroupNamePrefixes  # Array of prefixes to filter group names
    )

    # Authenticate to Microsoft Graph
    Connect-MicrosoftGraphAPI -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId

    $groupsWithMembers = @()  # Array to hold groups with their distinct members
    $UserCache = @{}  # Hashtable to cache user data (ensuring distinct users)

    # Retrieve all groups
    $groups = Get-MgGroup -All -Property 'id,displayName,groupTypes,description,mail' | Select-Object id, displayName, groupTypes, description, mail
    
    Write-Host "Total Groups Found: $($groups.Count)"

    # Filter groups based on name prefixes
    $filteredGroups = $groups | Where-Object {
        foreach ($prefix in $GroupNamePrefixes) {
            if ($_.displayName -like "$prefix*") {
                return $true
            }
        }
        return $false
    }

    Write-Host "Filtered Groups Count: $($filteredGroups.Count)"

    # Process each group sequentially
    foreach ($group in $filteredGroups) {
        Write-Host "Processing Group: $($group.displayName)"

        # Get distinct members for the group, including nested group members
        $distinctMembers = Get-AllDistinctGroupMembers -GroupId $group.Id -UserCache $UserCache -IsTopLevelGroup $true

        # Add the distinct members and their count
        $group | Add-Member -MemberType NoteProperty -Name "DistinctMembers" -Value ($distinctMembers | Sort-Object -Property Id -Unique) -Force
        $group | Add-Member -MemberType NoteProperty -Name "DistinctMemberCount" -Value ($distinctMembers | Sort-Object -Property Id -Unique).Count -Force

        # Add to the results list
        $groupsWithMembers += $group
    }

    # Disconnect the Graph session after completion
    Disconnect-MgGraph

    return $groupsWithMembers
}

# Export the functions
Export-ModuleMember -Function Connect-MicrosoftGraphAPI, Get-AllGroupsWithDistinctMembers
