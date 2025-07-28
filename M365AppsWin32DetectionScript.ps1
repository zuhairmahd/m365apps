#DetectionScript
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

$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\appLogs\M365AppsWin32DetectionScript.log"
$moduleName = "Office 365 detection"
Write-Log -LogFile $logFile -StartLogging
Write-Log -Message "Starting $moduleName" -LogFile $logFile -Module $moduleName 
Write-Host "Starting $moduleName" -ForegroundColor Cyan
$RegistryKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$M365Apps = "Microsoft 365"
$M365AppsCheck = $RegistryKeys | Get-ItemProperty | Where-Object { $_.DisplayName -match $M365Apps }
#log the above values.
Write-Log -Message "Registry Keys: $($RegistryKeys | Select-Object -ExpandProperty Name)" -LogFile $logFile -Module $moduleName
Write-Log -Message "M365Apps: $M365Apps" -LogFile $logFile -Module $moduleName
Write-Log -Message "M365AppsCheck: $($M365AppsCheck | Select-Object -ExpandProperty DisplayName)" -LogFile $logFile -Module $moduleName
if ($M365AppsCheck)
{
	Write-Log -Message "Microsoft 365 Apps Detected" -LogFile $logFile -Module $moduleName -LogLevel Information
	Write-Host "Microsoft 365 Apps Detected" -ForegroundColor Green
	Write-Log -LogFile $logFile -FinishLogging
	exit 0
}
else
{
	Write-Log -Message "Microsoft 365 Apps Not Detected" -LogFile $logFile -Module $moduleName -LogLevel Warning
	Write-Host "Microsoft 365 Apps Not Detected" -ForegroundColor Red
	Write-Log -LogFile $logFile -FinishLogging
	exit 1
}
Write-Log -LogFile $logFile -FinishLogging
