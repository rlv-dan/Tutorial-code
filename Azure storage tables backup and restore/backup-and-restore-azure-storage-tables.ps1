<#
    Simple script for backing up and restoring Azure Storage tables. Data 
    is save on your local disk.

    Prerequisites
        Download and install AzCopy: http://aka.ms/downloadazcopy
        
        More info about AzCopy:
        https://docs.microsoft.com/en-us/azure/storage/storage-use-azcopy

    Backing up tables
        1. Edit the config section
        2. Run Init-StorageAccountTablesBackup to create tables.csv
        3. Open tables.csv. Verify content and remove any table you don't want to backup.
        4. Run Backup-StorageAccountTables

    Restore tables
        1. First run Remove-StorageAccountTables
        2. Wait a while to let Azure remove your tables
        3. Run Restore-StorageAccountTables
#>


### Config ###

$backupDir = "c:\backup"
$AzCopyFolder = "C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy\"

$storageAccountName = "mystorageaccount"
$storageAccountKey = "rtyuiertyuqweasdfghjqwekliopasdfopzxcvbnmqiopasdfopwghjklzxcvbnm=="


### Functions ###

function Init-StorageAccountTablesBackup {
    if(!(Test-Path $backupDir)) {
        Write-Host ("Creating folder $backupDir") -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $backupDir
    }
    $ctx = New-AzureStorageContext $storageAccountName -StorageAccountKey $storageAccountKey
    $tables = Get-AzureStorageTable -Context $ctx
    $tables | Sort-Object -Property "CloudTable" | Select-Object CloudTable,Uri | Export-Csv "$backupDir\tables.csv" -Encoding UTF8
}

function Remove-StorageAccountTables {
    $tables = import-csv "$backupDir\tables.csv" -Encoding UTF8
    $ctx = New-AzureStorageContext $storageAccountName -StorageAccountKey $storageAccountKey
    foreach($t in $tables) {
        $tableName = $t.CloudTable
        Write-Host ("Removing table $tableName") -ForegroundColor Yellow
        Remove-AzureStorageTable –Name $tableName –Context $ctx
    }
}

function Backup-StorageAccountTables {
    $tables = import-csv "$backupDir\tables.csv" -Encoding UTF8
    cd $AzCopyFolder
    foreach($t in $tables) {
        $tableUrl = $t.Uri
        $tableName = $t.CloudTable
        Write-Host ("`nBacking up table $tableName") -ForegroundColor Yellow
        .\AzCopy.exe /Source:"$tableUrl" /Dest:"$backupDir" /SourceKey:"$storageAccountKey" /Manifest:"$tableName.manifest"
    }
}

function Restore-StorageAccountTables {
    $tables = import-csv "$backupDir\tables.csv" -Encoding UTF8
    cd $AzCopyFolder
    foreach($t in $tables) {
        $tableUrl = $t.Uri
        $tableName = $t.CloudTable
        Write-Host ("`nRestoring table $tableName") -ForegroundColor Yellow
        .\AzCopy.exe /Source:"$backupDir" /Dest:"$tableUrl" /DestKey:"$storageAccountKey" /Manifest:"$tableName.manifest" /EntityOperation:InsertOrReplace
    }
}
