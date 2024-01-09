if (-not (Get-Module -ListAvailable -Name SentinelARConverter)) {
    Write-Host "SentinelARConverter module not found."
    $userConsent = Read-Host "Do you want to install the SentinelARConverter module? (Y/N)"
    if ($userConsent -eq 'Y') {
        Install-Module -Name SentinelARConverter -Force
        Import-Module SentinelARConverter
    } else {
        Write-Host "Module installation declined. Exiting script."
        exit
    }
} else {
    Write-Host "SentinelARConverter module is already installed, awesome!"
}

# Prompt the user for the source root path
$sourceRoot = Read-Host "Please enter the source path to the \Azure-Sentinel\Solutions folder"
if (-not [System.IO.Directory]::Exists($sourceRoot)) {
    Write-Host "Source path does not exist. Exiting script."
    exit
}

# Prompt the user for the destination root path
$destinationRoot = Read-Host "Please enter the destination path to where to drop the converted rules"
if (-not [System.IO.Directory]::Exists($destinationRoot)) {
    # Optionally create the destination directory if it does not exist
    $createDestination = Read-Host "Destination path does not exist. Do you want to create it? (Y/N)"
    if ($createDestination -eq 'Y') {
        New-Item -ItemType Directory -Path $destinationRoot
        Write-Host "Created the destination directory."
    } else {
        Write-Host "Exiting script."
        exit
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "dd-MM-yyyy"

# Create the folder name with the timestamp
$folderName = "AR_Conversion_$currentDateTime"
$fullFolderPath = Join-Path -Path $destinationRoot -ChildPath $folderName

# Create the folder
New-Item -Path $fullFolderPath -ItemType Directory

$solutionnaam = @(
    "Azure Activity",
    "Azure Key Vault",
    "Microsoft 365",
    "Microsoft Defender for Cloud",
    "Microsoft Defender for Cloud Apps",
    "Microsoft Defender XDR",
    "AzureDevOpsAuditing",
    "Endpoint Threat Protection Essentials",
    "Microsoft Entra ID"
)

# Initialize an array to keep track of selections
$selections = @()
foreach ($solution in $solutionnaam) {
    $selections += $false
}

function DisplaySelections {
    Clear-Host
    Write-Host "Please select/deselect the solutions (Enter the number and press Enter):"
    Write-Host "0: [Select All]"
    for ($i = 0; $i -lt $solutionnaam.Length; $i++) {
        $selectionStatus = if ($selections[$i]) { "[x]" } else { "[ ]" }
        Write-Host "$($i+1): $selectionStatus $($solutionnaam[$i])"
    }
}

do {
    DisplaySelections

    $userinput = Read-Host "Enter the number (or 'done' to finish)"
    if ($userinput -eq 'done') { break }

    # Check if 'Select All' option was chosen
    if ($userinput -eq '0') {
        $selections = $selections | ForEach-Object { $true }
    } else {
        $selectedIndex = [int]$userinput - 1
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $solutionnaam.Length) {
            $selections[$selectedIndex] = -not $selections[$selectedIndex]
        }
    }

} while ($true)

$selectedSolutions = @()
for ($i = 0; $i -lt $solutionnaam.Length; $i++) {
    if ($selections[$i]) {
        $selectedSolutions += $solutionnaam[$i]
    }
}

Write-Host "You have selected the following solutions:"
$selectedSolutions | ForEach-Object { Write-Host $_ }

$successfulYamlToJsonConversions = 0
$failedConversions = 0

foreach ($solution in $selectedSolutions) {
    $folderPath = Join-Path $sourceRoot "$solution\Analytic Rules\"

    $yamlFileNames = Get-ChildItem -Path $folderPath -Filter "*.yaml" -Recurse | ForEach-Object { $_.FullName }

    foreach ($item in $yamlFileNames) {
        try {

            # Construct the destination file path
            $relativeFolderPath = (Split-Path $item).Substring($sourceRoot.Length)
            $destinationFolderPath = Join-Path $fullFolderPath $relativeFolderPath
            $destinationFileName = [System.IO.Path]::ChangeExtension((Split-Path $item -Leaf), ".json")
            $destinationFilePath = Join-Path $destinationFolderPath $destinationFileName

            # Ensure the destination folder exists
            if (-not (Test-Path $destinationFolderPath)) {
                New-Item -ItemType Directory -Path $destinationFolderPath -Force
            }

            # Convert and write the output directly to the destination path
            Get-Content "$item" | Convert-SentinelARYamlToArm -OutFile "$destinationFilePath"

            Write-Host "Converted and saved $item to $destinationFilePath"
            $successfulYamlToJsonConversions++
        } catch {
            Write-Host "Failed to convert $item"
            Write-Host "$_"
            $failedConversions++
        }
    }
}

Write-Host "Conversion Summary:"
Write-Host "Successfully converted YAML to JSON: $successfulYamlToJsonConversions"
Write-Host "Failed conversions: $failedConversions"