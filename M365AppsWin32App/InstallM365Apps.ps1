<#
.SYNOPSIS
  Script to install M365 Apps as a Win32 App

.DESCRIPTION
    Script to install Office as a Win32 App during Autopilot by downloading the latest Office setup exe from evergreen url
    Running Setup.exe from downloaded files with provided config.xml file. 

.EXAMPLE
    Without external XML (Requires $installationFile in the package)
    powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1
    With external XML (Requires XML to be provided by URL)  
    powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"

.NOTES
    Version:        1.2
    Author:         Jan Ketil Skanke
    Contact:        @JankeSkanke
    Creation Date:  01.07.2021
    Updated:        (2022-23-11)
    Version history:
        1.0.0 - (2022-23-10) Script released 
        1.1.0 - (2022-25-10) Added support for External URL as parameter 
        1.2.0 - (2022-23-11) Moved from ODT download to Evergreen url for setup.exe 
        1.2.1 - (2022-01-12) Adding function to validate signing on downloaded setup.exe
#>
#region parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$XMLUrl,
    [switch]$Uninstall
)
#endregion parameters
#Region Functions
function Write-Log()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [string]$Message,
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $true, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $true, ParameterSetName = 'FinishLogging')]
        [ValidateScript({
                $parentDir = Split-Path $_ -Parent
                if (-not (Test-Path $parentDir))
                {
                    try
                    {
                        New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                    }
                    catch
                    {
                        throw "Failed to create log directory: $_. Exception: $($_.Exception.Message)"
                    }
                }
                return $true
            })]
        [string]$LogFile,
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [string]$Module,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [switch]$WriteToConsole,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [ValidateSet("Verbose", "Debug", "Information", "Warning", "Error")]
        [string]$LogLevel = "Information",
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [switch]$CMTraceFormat,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [int]$MaxLogSizeMB = 10,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [switch]$PassThru,
        [Parameter(Mandatory = $true, ParameterSetName = 'StartLogging')]
        [switch]$StartLogging,
        [Parameter(Mandatory = $true, ParameterSetName = 'FinishLogging')]
        [switch]$FinishLogging,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [ValidateSet('Error', 'Warning', 'Information', 'Verbose', 'Debug')]
        [string]$MinimumLogLevel
    )
    try
    {
        # Use global minimum log level if not provided
        if (-not $MinimumLogLevel -and $Global:MinimumLogLevel)
        {
            $MinimumLogLevel = $Global:MinimumLogLevel
        }
        elseif (-not $MinimumLogLevel)
        {
            $MinimumLogLevel = 'Information'
        }
        
        # Define log level hierarchy (higher numbers = more detailed logging)
        $logLevelHierarchy = @{
            'Error'       = 1
            'Warning'     = 2
            'Information' = 3
            'Verbose'     = 4
            'Debug'       = 5
        }
        
        # Handle StartLogging and FinishLogging switches
        if ($StartLogging -or $FinishLogging)
        {
            # Set default values when using StartLogging or FinishLogging
            $Module = $MyInvocation.MyCommand.Name
            $LogLevel = "Information"
            
            # Create separator line
            $separatorLine = "=" * 80
            
            # Ensure log directory exists
            $logDir = Split-Path $LogFile -Parent
            if (-not (Test-Path $logDir))
            {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
            
            # Check for log rotation if file exists and is too large
            if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt ($MaxLogSizeMB * 1MB))
            {
                $archiveFile = $LogFile -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                Move-Item -Path $LogFile -Destination $archiveFile -Force
                Write-Verbose "Log file rotated to: $archiveFile"
            }
            
            if ($CMTraceFormat)
            {
                # For CMTrace format, still use the separator but in CMTrace format
                $cmTime = Get-Date -Format "HH:mm:ss.fff+000"
                $cmDate = Get-Date -Format "MM-dd-yyyy"
                $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
                $logEntry = "<![LOG[$separatorLine]LOG]!><time=`"$cmTime`" date=`"$cmDate`" component=`"$Module`" context=`"`" type=`"1`" thread=`"$thread`" file=`"`">"
            }
            else
            {
                # For standard format, just use the separator line without timestamp
                $logEntry = $separatorLine
            }
            
            # Use mutex for thread safety
            $mutexName = "LogMutex_" + ($LogFile -replace '[\\/:*?"<>|]', '_')
            $mutex = New-Object System.Threading.Mutex($false, $mutexName)
            
            try
            {
                $mutex.WaitOne() | Out-Null
                Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8 -Force
            }
            finally
            {
                $mutex.ReleaseMutex()
                $mutex.Dispose()
            }
            
            # Write to console
            if ($OutputToConsole)
            {
                Write-Host $separatorLine
            }
            return
        }
        
        # Check if this log entry should be written based on minimum log level
        # Only continue if the current log level meets or exceeds the minimum threshold
        if (-not ($StartLogging -or $FinishLogging))
        {
            $currentLogLevelValue = $logLevelHierarchy[$LogLevel]
            $minimumLogLevelValue = $logLevelHierarchy[$MinimumLogLevel]
            
            if ($currentLogLevelValue -gt $minimumLogLevelValue)
            {
                # Current log level is more detailed than the minimum, skip logging to file
                # But still write to console streams
                switch ($LogLevel)
                {
                    "Error"
                    {
                        if ($OutputToConsole)
                        {
                            Write-Error "[$Module] $Message" -ErrorAction SilentlyContinue 
                        }
                    }
                    "Warning"
                    {
                        if ($OutputToConsole)
                        {
                            Write-Warning "[$Module] $Message" 
                        }
                    }
                    "Verbose"
                    {
                        if ($OutputToConsole)
                        {
                            Write-Verbose "[$Module] $Message" 
                        }
                    }
                    "Debug"
                    {
                        if ($OutputToConsole)
                        {
                            Write-Debug "[$Module] $Message" 
                        }
                    }
                    default
                    {
                        # For Information level, we don't output to console in this case
                    }
                }
                return
            }
        }
        
        # Ensure log directory exists
        $logDir = Split-Path $LogFile -Parent
        if (-not (Test-Path $logDir))
        {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Check for log rotation if file exists and is too large
        if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt ($MaxLogSizeMB * 1MB))
        {
            $archiveFile = $LogFile -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Move-Item -Path $LogFile -Destination $archiveFile -Force
            Write-Verbose "Log file rotated to: $archiveFile"
        }
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
        $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)        
        if ($CMTraceFormat)
        {
            # True CMTrace format: 
            $cmTime = Get-Date -Format "HH:mm:ss.fff+000"
            $cmDate = Get-Date -Format "MM-dd-yyyy"
            $severity = switch ($LogLevel)
            {
                "Error"
                {
                    3 
                }
                "Warning"
                {
                    2 
                }
                default
                {
                    1 
                }
            }
            $logEntry = "<![LOG[$Message]LOG]!><time=`"$cmTime`" date=`"$cmDate`" component=`"$Module`" context=`"`" type=`"$severity`" thread=`"$thread`" file=`"`">"
        }
        else
        {
            # Enhanced standard format with thread ID
            $logEntry = "$timestamp [$LogLevel] [$Module] [Thread:$thread] [Context:$Context] $Message"
        }
        
        # Use mutex for thread safety in concurrent scenarios
        $mutexName = "LogMutex_" + ($LogFile -replace '[\\/:*?"<>|]', '_')
        $mutex = New-Object System.Threading.Mutex($false, $mutexName)
        
        try
        {
            $mutex.WaitOne() | Out-Null
            Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8 -Force
        }
        finally
        {
            $mutex.ReleaseMutex()
            $mutex.Dispose()
        }
        
        # Write to appropriate PowerShell stream based on log level
        switch ($LogLevel)
        {
            "Error"
            {
                if ($OutputToConsole)
                {
                    Write-Error "[$Module] $Message" -ErrorAction SilentlyContinue 
                }
            }
            "Warning"
            {
                if ($OutputToConsole)
                {
                    Write-Warning "[$Module] $Message" 
                }
            }
            "Verbose"
            {
                if ($OutputToConsole)
                {
                    Write-Verbose "[$Module] $Message" 
                }
            }
            "Debug"
            {
                if ($OutputToConsole)
                {
                    Write-Debug "[$Module] $Message" 
                }
            }
            default
            {
                if ($OutputToConsole)
                {
                    Write-Verbose "Logged: $logEntry" 
                }
            }
        }
        
        # Return log entry if PassThru is specified
        if ($PassThru)
        {
            return [PSCustomObject]@{
                Timestamp = $timestamp
                LogLevel  = $LogLevel
                Module    = $Module
                Message   = $Message
                Thread    = $thread
                LogFile   = $LogFile
                Entry     = $logEntry
            }
        }
    }
    catch
    {
        Write-Error "Failed to write to log file '$LogFile': $_"
        # Fallback to console output
        Write-Host "$timestamp [$LogLevel] [$Module] $Message" -ForegroundColor $(
            switch ($LogLevel)
            {
                "Error"
                {
                    "Red" 
                }
                "Warning"
                {
                    "Yellow" 
                }
                "Debug"
                {
                    "Cyan" 
                }
                default
                {
                    "White" 
                }
            }
        )
    }
}

function Test-OfficeIsInstalled()
{
    [CmdletBinding()]
    param()
    $RegistryKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $M365Apps = "Microsoft 365"
    $M365AppsCheck = $RegistryKeys | Get-ItemProperty | Where-Object { $_.DisplayName -match $M365Apps }
    if ($M365AppsCheck)
    {
        Write-Output "Microsoft 365 Apps Detected"
        return $true
    }
    else
    {
        Write-LogEntry -Value "Microsoft 365 Apps not detected" -Severity 2
        return $false
    }
}

function Start-DownloadFile()
{
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    begin
    {
        Write-Output "[ACTION] Starting download of $Name from $URL to $Path"
        # Logging download start
        Write-Log -Message ("Starting download of $Name from $URL to $Path") -LogFile $LogFile -Module $ModuleName -LogLevel Information
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    process
    {
        try
        {
            # Create path if it doesn't exist
            if (-not(Test-Path -Path $Path))
            {
                Write-Output "[ACTION] Path $Path does not exist, creating it"
                Write-Log -Message ("Path $Path does not exist, creating it") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
            }
            # Start download of file
            Write-Output "[ACTION] Downloading $Name to $Path"
            Write-Log -Message ("Downloading $Name to $Path") -LogFile $LogFile -Module $ModuleName -LogLevel Information
            $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
            Write-Output "[SUCCESS] Download of $Name completed successfully."
            Write-Log -Message ("Download of $Name completed successfully.") -LogFile $LogFile -Module $ModuleName -LogLevel Information
        }
        catch
        {
            Write-Output "[ERROR] Failed to download $($Name): $($_.Exception.Message)"
            Write-Log -Message ('Failed to download ' + $Name + ': ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
            throw
        }
    }
    end
    {
        Write-Output "[ACTION] Disposing WebClient for $Name"
        # Dispose of the WebClient object
        Write-Log -Message ("Disposing WebClient for $Name") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
        $WebClient.Dispose()
    }
}

function Invoke-FileCertVerification()
{
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )
    Write-Output "[ACTION] Verifying setup file signature for $FilePath"
    Write-Log -Message ("Verifying setup file signature for $FilePath") -LogFile $LogFile -Module $ModuleName -LogLevel Information
    try
    {
        $Cert = (Get-AuthenticodeSignature -FilePath $FilePath).SignerCertificate
        $CertStatus = (Get-AuthenticodeSignature -FilePath $FilePath).Status
        Write-Output "[INFO] Certificate status: $CertStatus"
        Write-Log -Message ("Certificate status: $CertStatus") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
        if ($Cert)
        {
            Write-Output "[INFO] Certificate subject: $($Cert.Subject)"
            Write-Output "[INFO] Certificate issuer: $($Cert.Issuer)"
            Write-Output "[INFO] Certificate valid from: $($Cert.NotBefore) to $($Cert.NotAfter)"
            Write-Log -Message ("Certificate subject: $($Cert.Subject)") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
            Write-Log -Message ("Certificate issuer: $($Cert.Issuer)") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
            Write-Log -Message ("Certificate valid from: $($Cert.NotBefore) to $($Cert.NotAfter)") -LogFile $LogFile -Module $ModuleName -LogLevel Debug
            if ($cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid")
            {
                Write-Output "[SUCCESS] Certificate is valid and signed by Microsoft"
                Write-Log -Message ("Certificate is valid and signed by Microsoft") -LogFile $LogFile -Module $ModuleName -LogLevel Information
                $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
                $chain.Build($cert) | Out-Null
                $RootCert = $chain.ChainElements | ForEach-Object {$_.Certificate} | Where-Object {$PSItem.Subject -match "CN=Microsoft Root"}
                if (-not [string ]::IsNullOrEmpty($RootCert))
                {
                    Write-Output "[SUCCESS] Certificate chain verified to Microsoft Root: $($RootCert.Subject)"
                    Write-Log -Message ("Certificate chain verified to Microsoft Root: $($RootCert.Subject)") -LogFile $LogFile -Module $ModuleName -LogLevel Information
                    $TrustedRoot = Get-ChildItem -Path "Cert:\\LocalMachine\\Root" -Recurse | Where-Object { $PSItem.Thumbprint -eq $RootCert.Thumbprint}
                    if (-not [string]::IsNullOrEmpty($TrustedRoot))
                    {
                        Write-Output "[SUCCESS] Verified setupfile signed by : $($Cert.Issuer)"
                        Write-Log -Message ("Verified setupfile signed by : $($Cert.Issuer)") -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        return $True
                    }
                    else
                    {
                        Write-Output "[ERROR] No trust found to root cert - aborting"
                        Write-Log -Message ("No trust found to root cert - aborting") -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        return $False
                    }
                }
                else
                {
                    Write-Output "[ERROR] Certificate chain not verified to Microsoft - aborting"
                    Write-Log -Message ("Certificate chain not verified to Microsoft - aborting") -LogFile $LogFile -Module $ModuleName -LogLevel Error
                    return $False
                }
            }
            else
            {
                Write-Output "[ERROR] Certificate not valid or not signed by Microsoft - aborting"
                Write-Log -Message ("Certificate not valid or not signed by Microsoft - aborting") -LogFile $LogFile -Module $ModuleName -LogLevel Error
                return $False
            }
        }
        else
        {
            Write-Output "[ERROR] Setup file not signed - aborting"
            Write-Log -Message ("Setup file not signed - aborting") -LogFile $LogFile -Module $ModuleName -LogLevel Error
            return $False
        }
    }
    catch
    {
        Write-Output "[ERROR] Exception during file certificate verification: $($_.Exception.Message)"
        Write-Log -Message ('Exception during file certificate verification: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
        return $False
    }
}
#Endregion Functions

#Region Initialisations
$installationPath = "$env:TEMP\odt"
# Define log file path and module name for logging
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\appLogs"
$LogFile = "$LogFolder\InstallM365Apps.log"
$ModuleName = 'InstallM365Apps.ps1'
$installationFile = "install.xml"
$uninstallationFile = "uninstall.xml"
#Endregion Initialisations

#Initate Install
# Log script start
Write-Log -LogFile $LogFile -StartLogging
Write-Log -Message "Script started. Initiating Office setup process." -LogFile $LogFile -Module $ModuleName -LogLevel Information
# Log parameter values
Write-Log -Message "Parameter XMLUrl: $XMLUrl" -LogFile $LogFile -Module $ModuleName -LogLevel Debug
# Console output for script start and parameter values
Write-Output "[INFO] Script started. Initiating Office setup process."
Write-Output "[INFO] Parameter XMLUrl: $XMLUrl"
Write-Output "[INFO] Log file: $LogFile"
Write-Output "[INFO] Module name: $ModuleName"
Write-Output "[INFO] Installation path: $installationPath"
Write-Output "Initiating Office setup process"
#Attempt Cleanup of SetupFolder
if (Test-Path $installationPath)
{
    Write-Output "[ACTION] Cleaning up previous installation files in $installationPath"
    Write-Log -Message "Cleaning up previous installation files in $installationPath" -LogFile $LogFile -Module $ModuleName -LogLevel Information
    try
    {
        Remove-Item -Path $installationPath -Recurse -Force -ErrorAction Stop
        Write-Output "[SUCCESS] Successfully cleaned up $installationPath"
        Write-Log -Message "Successfully cleaned up $installationPath" -LogFile $LogFile -Module $ModuleName -LogLevel Information
    }
    catch
    {
        Write-Output "[ERROR] Failed to clean up $($installationPath): $($_.Exception.Message)"
        Write-Log -Message ('Failed to clean up ' + $installationPath + ': ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
    }
}

# Create new setup folder
try
{
    Write-Output "[ACTION] Creating new setup folder at $installationPath\\OfficeSetup"
    $SetupFolder = (New-Item -ItemType "directory" -Path $installationPath -Name OfficeSetup -Force).FullName
    Write-Output "[SUCCESS] Created new setup folder at $SetupFolder"
    Write-Log -Message "Created new setup folder at $SetupFolder" -LogFile $LogFile -Module $ModuleName -LogLevel Information
}
catch
{
    Write-Output "[ERROR] Failed to create setup folder: $($_.Exception.Message)"
    Write-Log -Message ('Failed to create setup folder: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
    Write-Log -LogFile $LogFile -FinishLogging
    exit 1
}

try
{
    # Download latest Office setup.exe
    $SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-Output "[ACTION] Attempting to download latest Office setup executable from $SetupEverGreenURL"
    Write-Log -Message "Attempting to download latest Office setup executable from $SetupEverGreenURL" -LogFile $LogFile -Module $ModuleName -LogLevel Information
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    try
    {
        # Start install preparations
        $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
        if (-not (Test-Path $SetupFilePath))
        {
            Write-Output "[ERROR] Setup file not found at $SetupFilePath"
            Write-Log -Message "Error: Setup file not found at $SetupFilePath" -LogFile $LogFile -Module $ModuleName -LogLevel Error
            throw "Error: Setup file not found"
        }
        Write-Output "[SUCCESS] Setup file ready at $SetupFilePath"
        Write-Log -Message "Setup file ready at $SetupFilePath" -LogFile $LogFile -Module $ModuleName -LogLevel Information
        try
        {
            # Prepare Office Installation
            $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$($SetupFolder)\setup.exe").FileVersion
            Write-Output "[INFO] Office C2R Setup is running version $OfficeCR2Version"
            Write-Log -Message "Office C2R Setup is running version $OfficeCR2Version" -LogFile $LogFile -Module $ModuleName -LogLevel Information
            if (Invoke-FileCertVerification -FilePath $SetupFilePath)
            {
                Write-Output "[SUCCESS] Setup file signature verified, proceeding with installation"
                Write-Log -Message "Setup file signature verified, proceeding with installation" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                if ($XMLUrl)
                {
                    Write-Output "[ACTION] Attempting to download $installationFile from external URL: $XMLUrl"
                    Write-Log -Message "Attempting to download $installationFile from external URL: $XMLUrl" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                    try
                    {
                        Start-DownloadFile -URL $XMLURL -Path $SetupFolder -Name $installationFile
                        Write-Output "[SUCCESS] Downloading $installationFile from external URL completed"
                        Write-Log -Message "Downloading $installationFile from external URL completed" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        Start-DownloadFile -URL $XMLURL -Path $SetupFolder -Name $uninstallationFile
                        Write-Output "[SUCCESS] Downloading $uninstallationFile from external URL completed"
                        Write-Log -Message "Downloading $uninstallationFile from external URL completed" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                    }
                    catch
                    {
                        Write-Output "[ERROR] Downloading $installationFile from external URL failed: $($_.Exception.Message)"
                        Write-Log -Message ('Downloading $installationFile from external URL failed: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        Write-Log -Message "M365 Apps setup failed" -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        Write-Log -LogFile $LogFile -FinishLogging
                        exit 1
                    }
                }
                else
                {
                    Write-Output "[ACTION] Running with local $installationFile"
                    Write-Log -Message "Running with local $installationFile" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                    try
                    {
                        Copy-Item "$($PSScriptRoot)\$installationFile" $SetupFolder -Force -ErrorAction Stop
                        Copy-Item "$($PSScriptRoot)\$uninstallationFile" $SetupFolder -Force -ErrorAction Stop

                        Write-Output "[SUCCESS] Local Office Setup configuration file copied"
                        Write-Log -Message "Local Office Setup configuration file copied" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                    }
                    catch
                    {
                        Write-Output "[ERROR] Failed to copy local $($installationFile): $($_.Exception.Message)"
                        Write-Log -Message ('Failed to copy local $($installationFile): ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        Write-Log -LogFile $LogFile -FinishLogging
                        exit 1
                    }
                }
                # Starting Office setup with configuration file
                try
                {
                    if ($uninstall)
                    {
                        Write-Output "[ACTION] Starting M365 Apps Uninstall with Win32App method"
                        Write-Log -Message "Starting M365 Apps Uninstall with Win32App method" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        Write-Output "[ACTION] Checking whether Office is installed."
                        Write-Log -Message "Checking whether Office is installed." -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        if (-not (Test-OfficeIsInstalled))
                        {
                            Write-Output "[ERROR] Office is not installed, nothing to do."
                            Write-Log -Message "Office is not installed, nothing to do." -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        }
                        else
                        {
                            Write-Output "[ACTION] Office is installed, proceeding with uninstallation."
                            Write-Log -Message "Office is installed, proceeding with uninstallation." -LogFile $LogFile -Module $ModuleName -LogLevel Information
                            $null = Start-Process $SetupFilePath -ArgumentList "/configure $($SetupFolder)\$uninstallationFile" -Wait -PassThru -ErrorAction Stop
                            if (-not (Test-OfficeIsInstalled))
                            {
                                Write-Output "[Success] Office uninstall was successful."
                                Write-Log -Message "Office uninstall was successful." -LogFile $LogFile -Module $ModuleName -LogLevel Information
                            }
                            else
                            {
                                Write-Output "[ERROR] Office uninstall failed, please check the logs."
                                Write-Log -Message "Office uninstall failed, please check the logs." -LogFile $LogFile -Module $ModuleName -LogLevel Error
                                Write-Log -LogFile $LogFile -FinishLogging
                                exit 1
                            }
                        }
                    }
                    else 
                    {
                        Write-Output "[ACTION] Starting M365 Apps Install with Win32App method"
                        Write-Log -Message "Starting M365 Apps Install with Win32App method" -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        $null = Start-Process $SetupFilePath -ArgumentList "/configure $($SetupFolder)\$installationFile" -Wait -PassThru -ErrorAction Stop
                        if (test-OfficeIsInstalled)
                        {
                            Write-Output "[SUCCESS] M365 Apps installation was successful."
                            Write-Log -Message "M365 Apps installation was successful." -LogFile $LogFile -Module $ModuleName -LogLevel Information
                        }
                        else
                        {
                            Write-Output "[ERROR] M365 Apps installation failed, please check the logs."
                            Write-Log -Message "M365 Apps installation failed, please check the logs." -LogFile $LogFile -Module $ModuleName -LogLevel Error
                        }
                    }
                }
                catch
                {
                    Write-Output "[ERROR] Error running the M365 Apps install: $($_.Exception.Message)"
                    Write-Log -Message ('Error running the M365 Apps install: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
                }
            }
            else
            {
                Write-Output "[ERROR] Unable to verify setup file signature"
                Write-Log -Message "Error: Unable to verify setup file signature" -LogFile $LogFile -Module $ModuleName -LogLevel Error
                throw "Error: Unable to verify setup file signature"
            }
        }
        catch
        {
            Write-Output "[ERROR] Error preparing office installation: $($_.Exception.Message)"
            Write-Log -Message ('Error preparing office installation: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
        }
    }
    catch
    {
        Write-Output "[ERROR] Error finding office setup file: $($_.Exception.Message)"
        Write-Log -Message ('Error finding office setup file: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
    }
}
catch
{
    Write-Output "[ERROR] Error downloading office setup file: $($_.Exception.Message)"
    Write-Log -Message ('Error downloading office setup file: ' + $_.Exception.Message) -LogFile $LogFile -Module $ModuleName -LogLevel Error
}
# #Cleanup 
#cleaning up
if (Test-Path "$installationPath\OfficeSetup")
{
    Write-Output "[ACTION] Cleaning up setup files in $installationPath\OfficeSetup"
    Write-Log -Message "Cleaning up setup files in $installationPath\OfficeSetup" -LogFile $LogFile -Module $ModuleName -LogLevel Information
    Remove-Item -Path "$installationPath\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "[SUCCESS] Successfully cleaned up setup files in $installationPath\OfficeSetup"
    Write-Log -Message "Successfully cleaned up setup files in $installationPath\OfficeSetup" -LogFile $LogFile -Module $ModuleName -LogLevel Information
}

# Log script finish
Write-Output "[INFO] Script completed."
Write-Log -Message "Script completed." -LogFile $LogFile -Module $ModuleName -LogLevel Information
Write-Log -LogFile $LogFile -FinishLogging