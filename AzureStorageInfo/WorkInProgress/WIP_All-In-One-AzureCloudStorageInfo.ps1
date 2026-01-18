# Requires Az module and ImportExcel (Install-Module ImportExcel if needed)
# Run as: .\AzureFootprintReport.ps1 -OutputPath "C:\Temp\AzureFootprint.xlsx"

param (
    [string]$OutputPath = "AzureFootprint.xlsx"
)

# Install/Check modules (comment out if already done)
if (-not (Get-Module -ListAvailable -Name Az)) { Install-Module Az -Scope CurrentUser -Force }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) { Install-Module ImportExcel -Scope CurrentUser -Force }

Import-Module Az
Import-Module ImportExcel

# Connect (interactive or use service principal if automated)
Connect-AzAccount

$subscriptions = Get-AzSubscription

# Function to write to Excel
function WriteTo-Excel {
    param ([array]$data, [string]$sheetName)
    if ($data.Count -gt 0) {
        $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -Append -FreezeTopRow
    }
}

# Clear file if exists
if (Test-Path $OutputPath) { Remove-Item $OutputPath }

Write-Host "Starting Azure footprint report..."

# Blob/Storage Account Usage (using metrics for actual consumed)
$storageData = @()
foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
    $accounts = Get-AzStorageAccount

    foreach ($acct in $accounts) {
        try {
            $metric = Get-AzMetric -ResourceId $acct.Id -MetricName "UsedCapacity" -WarningAction SilentlyContinue
            $usedBytes = ($metric.Data | Select-Object -Last 1).Average  # Latest value
            $usedGB = if ($usedBytes) { [math]::Round($usedBytes / 1GB, 2) } else { 0 }

            $storageData += [PSCustomObject]@{
                Subscription   = $sub.Name
                StorageAccount = $acct.StorageAccountName
                ResourceGroup  = $acct.ResourceGroupName
                Location       = $acct.Location
                UsedGB         = $usedGB
                Kind           = $acct.Kind
            }

            if ($storageData.Count -ge 50) {
                WriteTo-Excel -data $storageData -sheetName "StorageAccounts"
                $storageData = @()
            }
        } catch {
            Write-Host "Skip storage $($acct.StorageAccountName): $_"
        }
    }
}
if ($storageData.Count -gt 0) { WriteTo-Excel -data $storageData -sheetName "StorageAccounts" }

# SQL Databases (used from metrics, max from object)
$dbData = @()
foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
    $servers = Get-AzSqlServer

    foreach ($server in $servers) {
        try {
            $dbs = Get-AzSqlDatabase -ResourceGroupName $server.ResourceGroupName -ServerName $server.ServerName | Where-Object { $_.DatabaseName -ne "master" }

            foreach ($db in $dbs) {
                $metric = Get-AzMetric -ResourceId $db.ResourceId -MetricName "storage" -WarningAction SilentlyContinue
                $usedBytes = ($metric.Data | Select-Object -Last 1).Maximum  # Often max recent for used
                $usedGB = if ($usedBytes) { [math]::Round($usedBytes / 1GB, 2) } else { 0 }
                $maxGB = [math]::Round($db.MaxSizeBytes / 1GB, 2)

                $dbData += [PSCustomObject]@{
                    Subscription   = $sub.Name
                    Server         = $server.ServerName
                    Database       = $db.DatabaseName
                    ResourceGroup  = $server.ResourceGroupName
                    Location       = $server.Location
                    UsedGB         = $usedGB
                    MaxGB          = $maxGB
                }

                if ($dbData.Count -ge 50) {
                    WriteTo-Excel -data $dbData -sheetName "Databases"
                    $dbData = @()
                }
            }
        } catch {
            Write-Host "Skip SQL server $($server.ServerName): $_"
        }
    }
}
if ($dbData.Count -gt 0) { WriteTo-Excel -data $dbData -sheetName "Databases" }

# VMs (config disk sizes, not actual consumed)
$vmData = @()
foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
    $vms = Get-AzVM

    foreach ($vm in $vms) {
        try {
            $vmStatus = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status
            $power = ($vmStatus.Statuses | Where-Object { $_.Code -like "*PowerState*" }).DisplayStatus

            if ($power -notlike "*deallocated*") {
                $osDiskGB = [math]::Round($vm.StorageProfile.OsDisk.DiskSizeGB, 2)
                $dataDisksGB = ($vm.StorageProfile.DataDisks | Measure-Object -Property DiskSizeGB -Sum).Sum
                $totalDiskGB = $osDiskGB + $dataDisksGB

                $vmData += [PSCustomObject]@{
                    Subscription   = $sub.Name
                    VMName         = $vm.Name
                    Size           = $vm.HardwareProfile.VmSize
                    PowerState     = $power
                    Location       = $vm.Location
                    TotalDiskGB    = [math]::Round($totalDiskGB, 2)
                }
            }

            if ($vmData.Count -ge 50) {
                WriteTo-Excel -data $vmData -sheetName "VMs"
                $vmData = @()
            }
        } catch {
            Write-Host "Skip VM $($vm.Name): $_"
        }
    }
}
if ($vmData.Count -gt 0) { WriteTo-Excel -data $vmData -sheetName "VMs" }

Disconnect-AzAccount
Write-Host "Done! Check $OutputPath for sheets: StorageAccounts, Databases, VMs"