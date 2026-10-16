# Version 1.00
# Created By: 
# Release Date:
# Revised By: 
# Revised Date:

#Requires -Modules @{ ModuleName="ActiveDirectory"; ModuleVersion="1.0.0.0"}
#Requires -Modules @{ ModuleName="ADSync"; ModuleVersion="1.0.0.0"}
#Requires -RunAsAdministrator

# Script must run on Entra ID connect server.

# functions
function Initialize-HeaderNames # remember to tweak this function if your spreadsheet uses different header names
{
    $script:upn = "Work_Email"
    $script:jobTitle = "Position"
    $script:department = "Department_Desc"
    $script:managerEmail = "Primary_Supervisor_Email"
    $script:country = "Work_Location_Country"
    $script:state = "Work_Location_State"
    $script:city = "Work_Location_City"
    $script:clinic = "Clinic_Desc"    
}

function Show-Introduction
{
    Write-Host ("This script takes an import from the HRIS and updates the following properties for the given users in active directory and then pushes it to Entra ID:`n" +
                "- Job Title`n" +
                "- Department`n" +
                "- Manager`n" +
                "- Country`n" +
                "- State`n" +
                "- City`n" +
                "- Office Location`n`n" +
                "Script must run on Entra ID Connect server.") -ForegroundColor "Cyan"

}

function TryGet-UserCSV
{
    do
    {
        $csvPath = Prompt-UserCSVPath
        $userImport = Import-CSV -Path $csvPath -Encoding "UTF8"
    }
    while (-not(Test-CSVHasCorrectHeaders $userImport))

    Write-Host "CSV was found and validated.`n" -ForegroundColor "Green"

    return $userImport
}

function Prompt-UserCSVPath
{
    do
    {
        $csvPath = Read-Host "Enter path to user CSV (i.e. C:\UserExport.csv)"
        $csvPath = $csvPath.Trim('"') # trim quotes from the path if they were entered

        if ($csvPath -notlike "*.csv")
        {
            Write-Warning "Please enter path to a CSV file."
            continue
        }

        $csvFile = Get-Item -LiteralPath $csvPath -ErrorAction "SilentlyContinue"
        if ($null -eq $csvFile)
        {
            Write-Warning "File not found."
        }
    }
    while ($null -eq $csvFile)    
    return $csvPath
}

function Test-CSVHasCorrectHeaders($importedCSV)
{
    if ($importedCSV.Count -eq 0)
    {
        Write-Warning "Provided CSV is empty."
        Read-Host "Press Enter to exit"
        exit
    }
    
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $upn))
    {
        Write-Warning "This CSV file is missing a header called $upn to represent the user principal name."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $jobTitle))
    {
        Write-Warning "This CSV file is missing a header called $jobTitle to represent the user's job title."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $department))
    {
        Write-Warning "This CSV file is missing a header called $department to represent the user's department."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $managerEmail))
    {
        Write-Warning "This CSV file is missing a header called $managerEmail to represent the manager's email."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $country))
    {
        Write-Warning "This CSV file is missing a header called $country to represent the user's country."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $state))
    {
        Write-Warning "This CSV file is missing a header called $state to represent the user's state."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $city))
    {
        Write-Warning "This CSV file is missing a header called $city to represent the user's city."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType "NoteProperty" -Name $clinic))
    {
        Write-Warning "This CSV file is missing a header called $clinic to represent the user's clinic"
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Read-Host "Please make corrections to CSV and press Enter to continue" | Out-Null
    }

    return $validCSV
}

function Update-OrgChart($userImport)
{
    Write-Host "Updating org charts..." -ForegroundColor "Cyan"
    $timeStamp = Get-Date -Format "yyyy-MM-dd-hh-mmtt"
    $logFilePath = "$PSScriptRoot\OrgChartChanges $timeStamp.csv"
    $index = 0
    $recordCount = $userImport.Count

    foreach ($userRecord in $userImport)
    {
        $index++
        Write-Progress -Activity "Updating users..." -Status "$index/$recordCount - $($userRecord.$UPN)"
        $changeMade = $false

        $logRecord = New-LogRecord

        if ([string]::IsNullOrWhiteSpace($userRecord.$upn))
        {
            $logRecord.Status = "Skipped: blank UPN"
            $logRecord | Export-CSV -Path $logFilePath -Append -NoTypeInformation
            continue
        }

        # Trim whitespace from all provided attributes        
        $userUpn = $userRecord.$upn.Trim()
        $newJobTitle = $userRecord.$jobTitle.Trim()
        $newDepartment = $userRecord.$department.Trim()
        $newManagerEmail = $userRecord.$managerEmail.Trim()
        $newCountry = $userRecord.$country.Trim()
        $newState = $userRecord.$state.Trim()
        $newCity = $userRecord.$city.Trim()
        $newClinic = $userRecord.$clinic.Trim()

        $logRecord.UPN = $userUpn

        # Look up user in AD
        try
        {
            $adUser = Get-ADUser -Filter {UserPrincipalName -eq $userUpn} -Properties @("title", "department", "manager", "co", "st", "l", "physicalDeliveryOfficeName") -ErrorAction "Stop"
        }
        catch
        {
            $errorRecord = $_
            $logRecord.Status = $errorRecord.ToString()
            $logRecord | Export-CSV -Path $logFilePath -Append -NoTypeInformation
            continue
        }
        if ($null -eq $adUser)
        {
            $logRecord.Status = "User not found in AD"
            $logRecord | Export-CSV -Path $logFilePath -Append -NoTypeInformation
            continue
        }

        $errors = New-Object System.Collections.Generic.List[string]

        # Get manager info
        if ($null -ne $adUser.manager)
        {
            try
            {
                $manager = Get-ADUser -Identity $adUser.manager -ErrorAction "Stop"
            }
            catch
            {
                $errorRecord = $_
                $errors.Add("Error getting manager: $errorRecord")
            }
            if ($null -eq $manager)
            {
                $errors.Add("Current manager not found in AD.")
            }
        }

        # Fill in 'before' values on log record
        $logRecord.'Title Before' = $adUser.title
        $logRecord.'Department Before' = $adUser.department
        if ($manager) { $logRecord.'Manager Before' = $manager.name }        
        $logRecord.'Country Before' = $adUser.co
        $logRecord.'State Before' = $adUser.st
        $logRecord.'City Before' = $adUser.l
        $logRecord.'Clinic Before' = $adUser.physicalDeliveryOfficeName

        # Fill in 'after' values on log record. They will be the same as before unless changed later.
        $logRecord.'Title After' = $adUser.title
        $logRecord.'Department After' = $adUser.department
        if ($manager) { $logRecord.'Manager After' = $manager.name }
        $logRecord.'Country After' = $adUser.co
        $logRecord.'State After' = $adUser.st
        $logRecord.'City After' = $adUser.l
        $logRecord.'Clinic After' = $adUser.physicalDeliveryOfficeName
                
        # Update job title
        if (-not [string]::IsNullOrWhiteSpace($newJobTitle))
        {
            if ($adUser.title -ne $newJobTitle) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -Title $newJobTitle -ErrorAction "Stop"
                    $logRecord.'Title After' = $newJobTitle
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating title: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided job title is blank.")
        }

        # Update department
        if (-not [string]::IsNullOrWhiteSpace($newDepartment))
        {
            if ($adUser.department -ne $newDepartment) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -Department $newDepartment -ErrorAction "Stop"
                    $logRecord.'Department After' = $newDepartment
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating department: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided department is blank.")
        }

        # Update manager
        if (-not [string]::IsNullOrWhiteSpace($newManagerEmail))
        {
            if ($manager.UserPrincipalName -ne $newManagerEmail) # check if there are changes to make
            {
                try
                {
                    $newManager = Get-ADUser -Filter { UserPrincipalName -eq $newManagerEmail} -ErrorAction "Stop"
                    if ($null -ne $manager)
                    {
                        Set-ADUser -Identity $adUser -Manager $newManager.objectGUID -ErrorAction "Stop"
                        $logRecord.'Manager After' = $newManager.name
                        $changeMade = $true
                    }
                    else
                    {
                        $errors.Add("Manager not found in AD. Manager unchanged.")
                    }
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating manager: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided manager is blank.")
        }

        # Update country
        if (-not [string]::IsNullOrWhiteSpace($newCountry))
        {
            if ($adUser.co -ne $newCountry) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -Replace @{co=$newCountry} -ErrorAction "Stop"
                    $logRecord.'Country After' = $newCountry
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating country: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided country is blank")
        }

        # Update state
        if (-not [string]::IsNullOrWhiteSpace($newState))
        {
            if ($adUser.st -ne $newState) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -State $newState -ErrorAction "Stop"
                    $logRecord.'State After' = $newState
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating state: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided state is blank")
        }

        # Update city
        if (-not [string]::IsNullOrWhiteSpace($newCity))
        {
            if ($adUser.l -ne $newCity) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -City $newCity -ErrorAction "Stop"
                    $logRecord.'City After' = $newCity
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating city: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided city is blank")
        }

        # Update office location
        if ( -not([string]::IsNullOrWhiteSpace($newClinic)) -and -not([string]::IsNullOrWhiteSpace($newState)) )
        {
            $newOfficeLocation = "$newState - $newClinic"
            if ($adUser.physicalDeliveryOfficeName -ne $newOfficeLocation) # check if there are changes to make
            {
                try
                {
                    Set-ADUser -Identity $adUser -Office $newOfficeLocation -ErrorAction "Stop"
                    $logRecord.'Clinic After' = $newOfficeLocation
                    $changeMade = $true
                }
                catch
                {
                    $errorRecord = $_
                    $errors.Add("Error updating office location: $errorRecord")
                }
            }
        }
        else
        {
            $errors.Add("Provided clinic or state is blank")
        }

        # Set status
        if ($errors.Count -eq 0) # if no errors
        {
            if ($changeMade)
            {
                $logRecord.Status = "Updated successfully"
            }
            else
            {
                $logRecord.Status = "No changes"
            }            
        }
        else
        {
            $errorMessage = $errors -Join ' -- '
            if ($changeMade)
            {
                $logRecord.Status = "Partial success with errors: $errorMessage" 
            }
            else
            {
                $logRecord.Status = "No changes due to errors: $errorMessage"
            }
        }
        
        # Export log record
        $logRecord | Export-CSV -Path $logFilePath -Append -NoTypeInformation
    }
    Write-Progress -Activity "Updating users..." -Completed
}

function New-LogRecord
{
    return [PSCustomObject]@{ 
    "UPN"               = ""
    "Status"            = ""
    "Title Before"      = ""
    "Title After"       = ""
    "Department Before" = ""
    "Department After"  = ""
    "Manager Before"    = ""
    "Manager After"     = ""
    "Country Before"    = ""
    "Country After"     = ""
    "State Before"      = ""
    "State After"       = ""
    "City Before"       = ""
    "City After"        = ""
    "Clinic Before"     = ""
    "Clinic After"      = ""
    }
}

function Sync-ToEntraID
{
    try 
    {
        $result = Start-ADSyncSyncCycle -PolicyType "Delta" -ErrorAction "Stop" # Syncs on-prem changes to Entra ID
    }
    catch 
    {
        $errorRecord = $_
        Write-Warning "Error starting sync to Entra ID: `n$errorRecord"
        return
    }

    if ($result.Result -ieq "Success")
    {
        Write-Host "Sync to Entra ID was started successfully." -ForegroundColor "Green"
    }
    else
    {
        Write-Warning "Issue starting sync to Entra ID. Result was $($result.Result)."
    }
}

# main
Initialize-HeaderNames
Show-Introduction
$userImport = TryGet-UserCSV
Read-Host "Press Enter to make the changes"
Update-OrgChart $userImport
Sync-ToEntraID
Read-Host "Press Enter to exit"