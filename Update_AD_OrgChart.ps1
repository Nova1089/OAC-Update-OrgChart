# Version 1.00
# Created By: Zak Paine
# Release Date: 10/15/2026
# Revised By: 
# Revised Date:

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

}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor "Cyan"
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor "Cyan"
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
            "1. Open Powershell as admin.`n" +
            "2. CD into script directory.`n" +
            "3. Run .\scriptname`n") -ForegroundColor "Red"
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryGet-UserCSV
{
    do
    {
        $csvPath = Prompt-UserCSVPath
        $userImport = Import-CSV -Path $csvPath
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

        $csvFile = Get-Item -Path $csvPath -ErrorAction SilentlyContinue
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
    Write-Host "Updating org charts..." -ForegroundColor "DarkCyan"
    $timeStamp = Get-Date -Format "yyyy-MM-dd-hh-mm"
    $logFilePath = "$PSScriptRoot\OrgChartChanges $timeStamp.csv"
    $totalRecordsProcessed = 0

    foreach ($user in $userImport)
    {
        $totalRecordsProcessed++
        Write-Progress -Activity "Updating users with new info..." -Status "$totalRecordsProcessed records processed."

        $userUpn = $user.$upn.Trim()
        $logRecord = New-LogRecord

        try
        {
            $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties @("title", "department", "manager", "co", "st", "l", "physicalDeliveryOfficeName")
        }
        catch
        {
            $errorRecord = $_
            $logRecord.Status = $errorRecord.ToString()
            continue
        }        

        # Fill in 'before' values on log record
        $logRecord.UPN = $userUpn
        $logRecord.'Title Before' = $adUser.title
        $logRecord.'Department Before' = $adUser.department
        $managerName = ($adUser.manager -split ',')[0] -replace '^CN=', '' # get the canonical name from the distinguished name
        $logRecord.'Manager Before' = $managerName
        $logRecord.'Country Before' = $adUser.co
        $logRecord.'State Before' = $adUser.start
        $logRecord.'City Before' = $adUser.l
        $logRecord.'Clinic Before' = $adUser.physicalDeliveryOfficeName

        # Update user and fill in log record
        try
        {
            Set-ADUser -Identity $adUser -Title $user.$jobTitle
            $logRecord.'Title After' = $user.$jobTitle
        }
        catch
        {
            $errorRecord = $_
            $logRecord.Status = $errorRecord.ToString()
        }

        try
        {

        }
        catch
        {

        }

        try
        {

        }
        catch
        {

        }
    }
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

# main
Initialize-HeaderNames
Show-Introduction
Use-Module "ActiveDirectory"
$userImport = TryGet-UserCSV
Read-Host "Press Enter to make the changes"
Update-OrgChart $userImport
Read-Host "Press Enter to exit"