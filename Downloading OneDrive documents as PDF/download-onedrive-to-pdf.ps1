
### This script download Word documents from OneDrive and saves as local PDF files
### Requires Microsoft Word and SharePoint Online Client Components SDK (https://www.microsoft.com/en-us/download/details.aspx?id=42038)

$SharePointClientDll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SharePoint Client Components\'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Location') + "ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path $SharePointClientDll 


function Get-OneDriveFiles {
    # Based on script from http://gsexdev.blogspot.com/2015/04/downloading-shared-file-from-onedrive.html
    param (
        $Tenant = "$( throw 'Tenant is a mandatory Parameter' )",
        $FileUrls = "$( throw 'FileUrls is a mandatory Parameter' )",
        $PSCredentials = "$( throw 'PSCredentials is a mandatory Parameter' )",
        $DownloadPath  = "$( throw 'DownloadPath is a mandatory Parameter' )"
    )
    process {
        $SpoCredentials =  New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($PSCredentials.UserName.ToString(),$PSCredentials.password) 
        $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($Tenant)
        $clientContext.Credentials = $SpoCredentials;
        $output = @()
        ForEach($url in $FileUrls) {
            $DownloadURI = New-Object System.Uri($url)
            $destPath = ($DownloadPath + [System.IO.Path]::GetFileName($DownloadURI.LocalPath))
            Write-Host "Downloading $url..." -NoNewline
            $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($clientContext, $DownloadURI.LocalPath);
            $fstream = New-Object System.IO.FileStream($destPath, [System.IO.FileMode]::Create);
            $fileInfo.Stream.CopyTo($fstream)
            $fstream.Flush()
            $fstream.Close()
            $output += $destPath
            Write-Host " Done!"
        }
    }
    end {
        $output
    }
}


function Convert-WordFilesToPdf {
 	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$True,
		           ValueFromPipeline=$True)]
		[string[]]$WordFiles,
        $SavePath  = "$( throw 'SavePath is a mandatory Parameter' )",
        [bool]$DeleteOriginal = $false
	)
    process {
	    $word_app = New-Object -ComObject Word.Application
        $outputFolder = Get-Item -LiteralPath $SavePath
	    foreach($f in $WordFiles) {
            try {
                Write-Host "Converting $f to PDF..." -NoNewline
                $file = Get-Item -LiteralPath $f
		        $document = $word_app.Documents.Open($file.FullName)
		        $pdf_filename = "$($SavePath)\$($file.BaseName).pdf"
		        $document.SaveAs([ref] $pdf_filename, [ref] 17)
		        $document.Close()
                if($DeleteOriginal) {
                    Remove-Item $file | Out-Null
                }
                Write-Host " Done!"
	        }
            catch {
            }
        }
	    $word_app.Quit()
    }
}


### Usage ###

# Prompt for credentials
$cred = Get-Credential

# If you don't want to enter your credentials every time you 
# run the script, and are ok with saving your password in the 
# script, use below instead of above promt
<#
$username = "your@email"
$password = "yourpassword"
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr
#>

# Configure what to do and where to put the files
$tenant = 'https://mytenant.sharepoint.com'
$docs = 'https://mytenant.sharepoint.com/sites/mysite/Shared%20documents/Doc1.docx',
        'https://mytenant.sharepoint.com/sites/mysite/Shared%20documents/Doc2.docx'
$downloadDocxTo = 'c:\'
$convertPdfTo = 'c:\'
$removeDocx = $true

# Run the cmdlets
Get-OneDriveFiles -Tenant $tenant -FileUrls $docs -PSCredentials $cred -DownloadPath $downloadDocxTo | Convert-WordFilesToPdf -SavePath $convertPdfTo -DeleteOriginal $removeDocx
