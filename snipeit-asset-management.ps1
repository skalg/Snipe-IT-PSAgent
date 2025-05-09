# Define the necessary variables
$SnipeItApiUrl = "https://your-snipe-it-instance/api/v1"
$SnipeItApiToken = "your_api_token"

# Static fields for asset creation
$status_id = 5  # Change this to the appropriate status ID for your assets
$fieldset_id = 1  # Change this to the appropriate fieldset ID for your models (Custom Fields)

# Function to load the necessary assembly for System.Web.HttpUtility
function Load-HttpUtilityAssembly {
    Add-Type -AssemblyName "System.Web"
}

# Function to check if Hyper-V is installed and list VMs
function Get-HyperVVMs {
    if (Get-Module -ListAvailable -Name "Hyper-V") {
        try {
            $vms = Get-VM | Select-Object -ExpandProperty Name
            if ($vms) {
                return $vms -join ", "
            } else {
                return ""
            }
        } catch {
            return ""
        }
    } else {
        return ""
    }
}

# Function to determine if the computer is a laptop or desktop
function Get-ComputerType {
    $battery = Get-WmiObject -Class Win32_Battery

    if ($battery) {
        return "Laptop"
    } else {
        return "Desktop"
    }
}

# Function to get the category ID based on computer type
function Get-CategoryId {
    $computerType = Get-ComputerType

    switch ($computerType) {
        "Laptop" { return 2 }
        "Desktop" { return 3 }
        default { return 3 }
    }
}

# Function to get the computer model
function Get-ComputerModel {
    $invalidModels = @(
        "Virtual Machine",
        "VMware Virtual Platform",
        "VM",
        "Parallels ARM Virtual Machine",
        $null
    )

    # Attempt to retrieve the manufacturer and model information
    $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
    $model = if ($manufacturer -match "Lenovo") {
        (Get-CimInstance -ClassName Win32_BIOS).Description
    } else {
        (Get-CimInstance -ClassName Win32_ComputerSystem).Model
    }

    # Verify that the model is valid
    if (-not $model) {
        Write-Warning "Model information is empty or null. Returning an empty string."
        return ""
    }

    if ($model -in $invalidModels) {
        Write-Warning "Model matches invalid list: '$model'. Returning an empty string."
        return ""
    }

    return $model
}

# Function to get the computer serial number
function Get-ComputerSerialNumber {
    $invalidSerials = @(
        "To Be Filled By O.E.M.",
        "Default_String",
        "INVALID"
    )

    $serialNumber = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber

    if ($serialNumber -in $invalidSerials) {
        return ""
    }

    return $serialNumber
}

# Function to get all MAC addresses of the computer
function Get-MacAddresses {
    $macAddresses = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.MACAddress -ne $null } | Select-Object -ExpandProperty MACAddress
    return $macAddresses -join ", "
}

# Function to get the RAM amount in GB
function Get-RAMAmount {
    $ramAmount = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
    return $ramAmount
}

# Function to get the CPU information
function Get-CPUInfo {
    $cpuInfo = (Get-WmiObject -Class Win32_Processor | Select-Object -First 1).Name
    return $cpuInfo
}

# Function to get the currently logged-on user
function Get-CurrentUser {
    $currentUser = "$env:USERDOMAIN\$env:USERNAME"
    return $currentUser
}

# Function to get the OS information
function Get-OSInfo {
    $osInfo = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    
    # Remove any non-alphanumeric characters and spaces
    $osInfo = ($osInfo -replace '[^\w\s]', '').Trim()
    
    # Replace non-breaking spaces (U+00A0) with a normal space (U+0020)
    $osInfo = $osInfo -replace '\u00A0', ' '

    # Trim leading and trailing spaces
    $osInfo = $osInfo.Trim()
    
    return $osInfo
}

# Function to get the Windows version
function Get-WindowsVersion {
    $windowsVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
    return $windowsVersion
}

# Function to get the build number
function Get-BuildNumber {
    $buildNumber = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
    return $buildNumber
}

# Function to get the kernel version
function Get-KernelVersion {
    $kernelVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version
    return $kernelVersion
}

# Function to get the current active IP address
function Get-ActiveIPAddress {
    $ipAddress = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true } | Select-Object -ExpandProperty IPAddress)[0]
    return $ipAddress
}

# Function to get storage type (SSD or HDD) and capacity
function Get-StorageInfo {
    $physicalDisks = Get-PhysicalDisk
    $storageInfo = @()
    foreach ($disk in $physicalDisks) {
        $type = if ($disk.MediaType -eq 'Unspecified' -or $disk.MediaType -eq $null) { 
            'Unknown' 
        } else { 
            $disk.MediaType 
        }
        $size = [math]::Round($disk.Size / 1GB, 2)
        $storageInfo += [PSCustomObject]@{
            Type = $type
            Capacity = "$size GB"
        }
    }
    return $storageInfo
}

# Gather information for custom fields
function Get-CustomFields {
    try {
        # Gather data for each custom dbfields
        #$macAddresses = Get-MacAddresses
        #$ramAmount = Get-RAMAmount
        #$cpuInfo = Get-CPUInfo
        #$currentUser = Get-CurrentUser
        #$osInfo = Get-OSInfo
        #$windowsVersion = Get-WindowsVersion
        #$buildNumber = Get-BuildNumber
        #$kernelVersion = Get-KernelVersion
        #$ipAddress = Get-ActiveIPAddress
        #$storageInfo = Get-StorageInfo
        #$hyperVVMs = Get-HyperVVMs
        #$storageType = ($storageInfo | ForEach-Object { $_.Type }) -join ", "
        #$storageCapacity = ($storageInfo | ForEach-Object { $_.Capacity }) -join ", "

        # Add your custom dbfield names : https://snipe-it.readme.io/reference/updating-custom-fields
        $dbFields = @{
            #"_snipeit_adresse_mac_1"   = if ($macAddresses) { $macAddresses } else { "" }
            #"_snipeit_ram_5"           = if ($ramAmount) { $ramAmount } else { "" }
            #"_snipeit_cpu_6"           = if ($cpuInfo) { $cpuInfo } else { "" }
            #"_snipeit_utilisateur_11"  = if ($currentUser) { $currentUser } else { "" }
            #"_snipeit_os_14"           = if ($osInfo) { $osInfo } else { "" }
            #"_snipeit_version_41"      = if ($windowsVersion) { $windowsVersion } else { "" }
            #"_snipeit_build_43"        = if ($buildNumber) { $buildNumber } else { "" }
            #"_snipeit_kernel_42"       = if ($kernelVersion) { $kernelVersion } else { "" }
            #"_snipeit_adresse_ipv4_18" = if ($ipAddress) { $ipAddress } else { "" }
            #"_snipeit_type_stockage_7" = if ($storageType) { $storageType } else { "" }
            #"_snipeit_capacitac_stockage_8" = if ($storageCapacity) { $storageCapacity } else { "" }
            #"_snipeit_vm_28"           = if ($hyperVVMs) { $hyperVVMs } else { "" }
        }
        return $dbFields
    } catch {
        Write-Error "An error occurred while gathering custom fields: $_"
        return @{}
    }
}

# Function to search for a model in Snipe-IT
function Search-ModelInSnipeIt {
    param (
        [string]$ModelName
    )

    if (-not $ModelName -or $ModelName -eq "") {
        Write-Warning "ModelName is null or empty. Cannot search for a model."
        return $null
    }

    Add-Type -AssemblyName "System.Web"
    $encodedModelName = [System.Web.HttpUtility]::UrlEncode($ModelName)

    $url = "$SnipeItApiUrl/models?limit=50&offset=0&search=$encodedModelName&sort=created_at&order=asc"
    $headers = @{
        "Authorization" = "Bearer $SnipeItApiToken"
        "accept"        = "application/json"
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get

        if (-not $response -or -not $response.total -or $response.total -eq 0) {
            Write-Warning "No models found for ModelName: '$ModelName'."
            return $null
        }

        # Check the response for a matching model
        foreach ($model in $response.rows) {
            if ($model.name -eq $ModelName) {
                Write-Output "Model found with ID: $($model.id)"
                return $model.id
            }
        }

        # If no exact match is found
        Write-Warning "No exact match found for ModelName: '$ModelName'."
        return $null
    } catch {
        # Handle errors during the API call
        Write-Error "An error occurred during the API request: $_"
        Write-Output "DEBUG: URL: $url"
        return $null
    }
}

# Function to create a model in Snipe-IT
function Create-ModelInSnipeIt {
    param (
        [string]$ModelName,
        [int]$CategoryId
    )

    # Validate input
    if (-not $ModelName -or $ModelName -eq "") {
        Write-Warning "ModelName is null or empty. Cannot create a model."
        return $null
    }
    if (-not $CategoryId -or $CategoryId -le 0) {
        Write-Warning "Invalid CategoryId provided. Cannot create a model."
        return $null
    }

    $url = "$SnipeItApiUrl/models"
    $headers = @{
        "Authorization" = "Bearer $SnipeItApiToken"
        "accept"        = "application/json"
        "content-type"  = "application/json"
    }

    $body = @{
        category_id = $CategoryId
        name        = $ModelName
    }

    # Conditionally add fieldset_id if available
    if ($fieldset_id -ne $null -and $fieldset_id -ne 0) {
        $body.fieldset_id = $fieldset_id
    }

    $bodyJson = $body | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $bodyJson

        # Validate the response payload
        if ($response -and $response.payload -and $response.payload.id) {
            return $response.payload.id
        } else {
            Write-Warning "Model creation response is missing expected fields. Response: $($response | ConvertTo-Json -Depth 10)"
            return $null
        }
    } catch {
        # Handle errors during the API call
        Write-Error "An error occurred during model creation: $_"
        Write-Output "DEBUG: URL: $url"
        Write-Output "DEBUG: Body: $bodyJson"
        return $null
    }
}

# Function to search for an asset in Snipe-IT
function Search-AssetInSnipeIt {
    param (
        [string]$SerialNumber
    )

    $encodedSerialNumber = $SerialNumber
    $url = "$SnipeItApiUrl/hardware?limit=50&offset=0&search=$encodedSerialNumber&sort=created_at&order=asc"
    $headers = @{
        "Authorization" = "Bearer $SnipeItApiToken"
        "accept"        = "application/json"
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        
        if ($response.total -gt 0) {
            foreach ($asset in $response.rows) {
                if ($asset.serial -eq $SerialNumber) {
                    return $asset
                }
            }
        }
    } catch {
        Write-Output "Error during asset search: $_"
    }

    return $null
}

# Function to create an asset in Snipe-IT
function Create-AssetInSnipeIt {
    param (
        [string]$ModelId,
        [string]$SerialNumber,
        [string]$AssetName,
        [hashtable]$CustomFields
    )

    # Validate inputs
    if (-not $ModelId -or $ModelId -eq "") {
        Write-Warning "ModelId is null or empty. Cannot create an asset."
        return $null
    }
    if (-not $SerialNumber -or $SerialNumber -eq "") {
        Write-Warning "SerialNumber is null or empty. Cannot create an asset."
        return $null
    }
    if (-not $AssetName -or $AssetName -eq "") {
        Write-Warning "AssetName is null or empty. Cannot create an asset."
        return $null
    }

    $url = "$SnipeItApiUrl/hardware"
    $headers = @{
        "Authorization" = "Bearer $SnipeItApiToken"
        "accept"        = "application/json"
        "content-type"  = "application/json"
    }

    $body = @{
        model_id  = $ModelId
        serial    = $SerialNumber
        name      = $AssetName
        status_id = $status_id
    } + $CustomFields | ConvertTo-Json -Depth 10

    try {
        # Make the API call
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body

        if ($response -and $response.payload -and $response.payload.id) {
            return $response.payload.id
        } else {
            Write-Warning "Asset creation response is missing expected fields. Response: $($response | ConvertTo-Json -Depth 10)"
            return $null
        }
    } catch {
        # Handle errors during the API call
        Write-Error "An error occurred during asset creation: $_"
        Write-Output "DEBUG: URL: $url"
        Write-Output "DEBUG: Body: $body"
        return $null
    }
}

# Function to update an asset in Snipe-IT
function Update-AssetInSnipeIt {
    param (
        [string]$AssetId,
        [string]$AssetName,
        [hashtable]$CustomFields
    )

    # Validate inputs
    if (-not $AssetId -or $AssetId -eq "") {
        Write-Warning "AssetId is null or empty. Cannot update an asset."
        return $null
    }
    if (-not $AssetName -or $AssetName -eq "") {
        Write-Warning "AssetName is null or empty. Cannot update an asset."
        return $null
    }

    $url = "$SnipeItApiUrl/hardware/$AssetId"
    $headers = @{
        "Authorization" = "Bearer $SnipeItApiToken"
        "accept"        = "application/json"
        "content-type"  = "application/json"
    }

    $body = @{
        name = $AssetName
    } + $CustomFields | ConvertTo-Json -Depth 10

    try {
        # Make the API call
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body

        if ($response -and $response.payload -and $response.payload.id) {
            #Write-Output "Asset updated successfully with ID: $($response.payload.id)"
            return $response.payload.id
        } else {
            Write-Warning "Asset update response is missing expected fields. Response: $($response | ConvertTo-Json -Depth 10)"
            return $null
        }
    } catch {
        # Handle errors during the API call
        Write-Error "An error occurred during asset update: $_"
        Write-Output "DEBUG: URL: $url"
        Write-Output "DEBUG: Body: $body"
        return $null
    }
}

# Main script logic
$computerModel = Get-ComputerModel
$serialNumber = Get-ComputerSerialNumber

if ($serialNumber) {
    $asset = Search-AssetInSnipeIt -SerialNumber $serialNumber

    $assetName = $env:COMPUTERNAME
    $customFields = Get-CustomFields

    if ($asset) {
        $assetId = $asset.id
        $updateRequired = $false

        if ($asset.name -ne $assetName) {
            Write-Output "Asset name requires update: '$($asset.name)' -> '$assetName'"
            $updateRequired = $true
        }

        foreach ($key in $customFields.Keys) {
            foreach ($field in $asset.custom_fields.PSObject.Properties) {
                if ($field.Value.field -eq $key -and $field.Value.value -ne $customFields[$key]) {
                    Write-Output "Custom field '$key' requires update: '$($field.Value.value)' -> '$($customFields[$key])'"
                    $updateRequired = $true
                    break
                }
            }
        }

        if ($updateRequired) {
            #Write-Output "DEBUG : $customFields"
            $updatedAssetId = Update-AssetInSnipeIt -AssetId $assetId -AssetName $assetName -CustomFields $customFields
            Write-Output "Asset updated with ID: $updatedAssetId"
        } else {
            Write-Output "No update required for asset with ID: $assetId"
        }
    } else {
        if ($computerModel) {
            $modelId = Search-ModelInSnipeIt -ModelName $computerModel

        if (-not $modelId) {
            $categoryId = Get-CategoryId
            $modelId = Create-ModelInSnipeIt -ModelName $computerModel -CategoryId $categoryId
        }

        $newAssetId = Create-AssetInSnipeIt -ModelId $modelId -SerialNumber $serialNumber -AssetName $assetName -CustomFields $customFields
        Write-Output "New Asset ID: $newAssetId"
        } else {
            Write-Output "Computer model could not be determined."
        }
    }
} else {
    Write-Output "No serial number found on this computer."
}
