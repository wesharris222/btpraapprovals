
# Azure Bot Deployment Script

# Set execution policy at the start
try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
} catch {
    Write-Error "Failed to set execution policy. Please run PowerShell as Administrator."
    exit 1
}

# Check PowerShell version and prompt for installation if needed
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "PowerShell 7 or higher is required. Currently running PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    $install = Read-Host "Would you like to install PowerShell 7 now? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            # Download PowerShell 7 installer
            Write-Host "Downloading PowerShell 7..." -ForegroundColor Yellow
            $url = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.1/PowerShell-7.4.1-win-x64.msi"
            $outPath = "$env:TEMP\PowerShell-7.msi"
            Invoke-WebRequest -Uri $url -OutFile $outPath

            # Install PowerShell 7
            Write-Host "Installing PowerShell 7..." -ForegroundColor Yellow
            Start-Process msiexec.exe -Wait -ArgumentList "/i `"$outPath`" /quiet"
            
            Write-Host "PowerShell 7 installed successfully. Please close this window and run the script using PowerShell 7 (pwsh.exe)" -ForegroundColor Green
            Write-Host "You can find PowerShell 7 by searching for 'pwsh' in the Start menu." -ForegroundColor Green
        } catch {
            Write-Error "Failed to install PowerShell 7: $_"
            Write-Host "Please install PowerShell 7 manually from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
        }
        exit 1
    } else {
        Write-Error "Script cannot continue without PowerShell 7. Please install it and try again."
        exit 1
    }
}

# Check if running in PowerShell Admin mode
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator. Please restart PowerShell 7 as Administrator."
    exit 1
}

Write-Host "PowerShell environment check passed. Proceeding with deployment..." -ForegroundColor Green

#region Required Variables
# =============================================
# CONFIGURATION: Edit these variables before running the script
# =============================================

# Bot and Azure Resource Settings
$BotName = ""                            # Name for the bot (will be used for various resources)
$ResourceGroupName = ""                  # Azure Resource Group name
$Location = "westus"                     # Azure region (e.g., 'eastus') do not change, as for some reason Bots can only be deployed to certain regions
$SubscriptionId = ""                     # Azure Subscription ID
$TenantId = ""                          # Azure AD Tenant ID

# Repository Settings
$GitRepoUrl = "https://github.com/wesharris222/btpraapprovals"  # URL to the git repository - dont change
$LocalRepoPath = ""                      # Local path where repo will be cloned (e.g., "C:\Projects\btpmapprovalsbot")
# PRA Settings
# $PraEndpointUrl = ""  # PRA endpoint URL for approvals <https://<prasitename>/api/endpoint_approval>
# Ensure local repository path exists
if (-not (Test-Path $LocalRepoPath)) {
    Write-Host "Creating local repository directory: $LocalRepoPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $LocalRepoPath -Force | Out-Null
}

# BeyondTrust Settings
#$BeyondTrustBaseUrl = ""                # Base URL for BeyondTrust API (https://xxxx-services.pm.beyondtrustcloud.com)
#$BeyondTrustClientId = ""               # BeyondTrust Client ID
#$BeyondTrustClientSecret = ""           # BeyondTrust Client Secret

# Office 365 Settings
$office365ConnectionName = "$BotName-office365" #leave this static
$monitoredEmailAddress = ""  # this is the inbox to receive the email

# Teams Settings
#$TeamsChannelName = "Bots Mansion"                  # Name of Teams channel to create - not used as of now.  will create manually

# Validate required variables are set
$requiredVariables = @(
    @{Name='BotName';Value=$BotName},
    @{Name='ResourceGroupName';Value=$ResourceGroupName},
    @{Name='SubscriptionId';Value=$SubscriptionId},
    @{Name='TenantId';Value=$TenantId},
    @{Name='LocalRepoPath';Value=$LocalRepoPath},
	@{Name='office365ConnectionName';Value=$office365ConnectionName},
	@{Name='monitoredEmailAddress';Value=$monitoredEmailAddress}
    
)

foreach ($var in $requiredVariables) {
    if ([string]::IsNullOrWhiteSpace($var.Value)) {
        Write-Error "Required variable '$($var.Name)' is not set. Please edit the script and set all required variables."
        exit 1
    }
}
#endregion

#region Check and Install Dependencies
function Install-RequiredModules {
    Write-Host "`n=== Starting dependency checks and installation ===" -ForegroundColor Cyan
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Beginning dependency verification..." -ForegroundColor Yellow

    Write-Host "`n=== Checking PowerShell Modules ===" -ForegroundColor Cyan
    # Check for required PowerShell modules
    $requiredModules = @(
        'Az',
        'Microsoft.Graph'
    )
    
    Write-Host "Required modules to check: $($requiredModules -join ', ')" -ForegroundColor Yellow
    
    foreach ($module in $requiredModules) {
        Write-Host "Importing module: $module... (This may take a few minutes for the Az module)" -ForegroundColor Yellow
        try {
            Import-Module -Name $module -Force
            Write-Host "Successfully imported $module" -ForegroundColor Green
        } catch {
            Write-Host "Error importing $module : $_" -ForegroundColor Red
            throw
        }
    }

    Write-Host "`n=== Checking Node.js and npm ===" -ForegroundColor Cyan
    try {
        $nodeVersion = node -v
        $npmVersion = npm -v
        Write-Host "Found Node.js version: $nodeVersion" -ForegroundColor Green
        Write-Host "Found npm version: $npmVersion" -ForegroundColor Green
    } catch {
        Write-Host "Node.js is not installed or not in PATH" -ForegroundColor Red
        $installNode = Read-Host "Would you like to install Node.js now? (Y/N)"
        
        if ($installNode -eq 'Y' -or $installNode -eq 'y') {
            try {
                # Download and install Node.js LTS
                $nodeUrl = "https://nodejs.org/dist/v20.11.1/node-v20.11.1-x64.msi"
                $nodeInstaller = "$env:TEMP\node-installer.msi"
                Write-Host "Downloading Node.js..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $nodeUrl -OutFile $nodeInstaller
                Write-Host "Installing Node.js..." -ForegroundColor Yellow
                Start-Process msiexec.exe -Wait -ArgumentList "/i `"$nodeInstaller`" /quiet /norestart"
                Remove-Item $nodeInstaller -Force
                
                # Refresh PATH
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
                
                Write-Host "Node.js installed successfully. Please restart PowerShell and run this script again." -ForegroundColor Green
                exit
            } catch {
                Write-Error "Failed to install Node.js: $_"
                Write-Host "Please install Node.js manually from: https://nodejs.org/" -ForegroundColor Yellow
            }
            exit 1
        } else {
            Write-Error "Node.js is required for this deployment. Please install from: https://nodejs.org/"
            exit 1
        }
    }

    Write-Host "`n=== Checking Azure CLI ===" -ForegroundColor Cyan
    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        Write-Host "Azure CLI not found. Installing..." -ForegroundColor Yellow
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Beginning Azure CLI installation..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi
            Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet /norestart'
            Remove-Item .\AzureCLI.msi -Force -ErrorAction SilentlyContinue
            
            # Add Azure CLI to PATH for current session
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
            
            Write-Host "Azure CLI installed successfully. You may need to restart PowerShell." -ForegroundColor Green
            exit 1
        } catch {
            Write-Error "Failed to install Azure CLI: $_"
            exit 1
        }
    }

    Write-Host "`n=== Checking Git Installation ===" -ForegroundColor Cyan
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Host "Git not found. Installing..." -ForegroundColor Yellow
        $installGit = Read-Host "Would you like to install Git now? (Y/N)"
        
        if ($installGit -eq 'Y' -or $installGit -eq 'y') {
            try {
                # Download Git installer
                Write-Host "Downloading Git for Windows..." -ForegroundColor Yellow
                $gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.43.0.windows.1/Git-2.43.0-64-bit.exe"
                $gitInstaller = "$env:TEMP\GitInstaller.exe"
                Invoke-WebRequest -Uri $gitUrl -OutFile $gitInstaller
                
                # Install Git silently
                Write-Host "Installing Git..." -ForegroundColor Yellow
                Start-Process -FilePath $gitInstaller -ArgumentList "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS" -Wait
                Remove-Item $gitInstaller -Force
                
                # Refresh PATH
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
                
                # Verify installation
                if (Get-Command git -ErrorAction SilentlyContinue) {
                    $gitVersion = git --version
                    Write-Host "Git installed successfully: $gitVersion" -ForegroundColor Green
                } else {
                    throw "Git installation completed but git command is not available. Please restart PowerShell and run the script again."
                }
            } catch {
                Write-Error "Failed to install Git: $_"
                Write-Host "Please install Git manually from: https://git-scm.com/" -ForegroundColor Yellow
                exit 1
            }
        } else {
            Write-Error "Git is required for this deployment. Please install Git and try again."
            exit 1
        }
    } else {
        $gitVersion = git --version
        Write-Host "Git is installed: $gitVersion" -ForegroundColor Green
    }

    Write-Host "`n=== Dependency checks completed ===" -ForegroundColor Cyan
}

function Connect-ToAzure {
    Write-Host "Clearing any existing Azure sessions..." -ForegroundColor Yellow
    az logout
    
    Write-Host "`nConnecting to Azure..." -ForegroundColor Yellow
    Write-Host "Please review the subscription list below and make a selection if prompted." -ForegroundColor Yellow
    Write-Host "Target subscription ID: $SubscriptionId" -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to cancel if you don't see your subscription.`n" -ForegroundColor Yellow
    
    az login
    $confirmation = Read-Host "`nPress Enter after you've selected your subscription"
    
    Write-Host "`nSetting subscription to: $SubscriptionId" -ForegroundColor Yellow
    az account set --subscription $SubscriptionId
    
    # Verify subscription was set correctly
    $currentSub = (az account show --output json | ConvertFrom-Json).id
    if ($currentSub -eq $SubscriptionId) {
        Write-Host "Successfully set subscription to: $SubscriptionId" -ForegroundColor Green
    } else {
        Write-Error "Failed to set subscription. Please verify your subscription ID"
        exit 1
    }
}

function New-AzureResources {
   Write-Host "`n=== Creating Azure Resources ===" -ForegroundColor Cyan

   try {
       # Create Resource Group if it doesn't exist
       Write-Host "Creating Resource Group: $ResourceGroupName" -ForegroundColor Yellow
       az group create --name $ResourceGroupName --location $Location

       # Check if app registration already exists
       Write-Host "Checking for existing app registration..." -ForegroundColor Yellow
       $existingApp = az ad app list --display-name $BotName --query "[0]" | ConvertFrom-Json

       if ($existingApp) {
           Write-Host "Found existing app registration with ID: $($existingApp.appId)" -ForegroundColor Yellow
           $appId = $existingApp.appId

           # Clean up existing credentials
           Write-Host "Cleaning up existing credentials..." -ForegroundColor Yellow
           $credentials = az ad app credential list --id $appId | ConvertFrom-Json
           foreach ($cred in $credentials) {
               az ad app credential delete --id $appId --key-id $cred.keyId
           }

           # Update existing app registration
           Write-Host "Updating existing app registration..." -ForegroundColor Yellow
           az ad app update --id $appId --display-name $BotName --sign-in-audience "AzureADMultipleOrgs"
       } else {
           Write-Host "Creating new Azure AD application..." -ForegroundColor Yellow
           $appRegistration = az ad app create --display-name $BotName --sign-in-audience "AzureADMultipleOrgs" | ConvertFrom-Json
           $appId = $appRegistration.appId
       }

       # Generate new client secret
       Write-Host "Generating new client secret..." -ForegroundColor Yellow
       $secret = az ad app credential reset `
           --id $appId `
           --display-name "BotSecret" | ConvertFrom-Json

       # Create Storage Account (ensure name is valid)
       Write-Host "Creating Storage Account..." -ForegroundColor Yellow
       $storageAccountName = ($BotName.ToLower() -replace '[^a-z0-9]', '') + "stor"
       if ($storageAccountName.Length -gt 24) {
           $storageAccountName = $storageAccountName.Substring(0, 24)
       }
       
       $storageAccount = az storage account create `
           --name $storageAccountName `
           --resource-group $ResourceGroupName `
           --location $Location `
           --sku Standard_LRS `
           --output json | ConvertFrom-Json

       # Retrieve storage connection string
       $storageConnString = (az storage account show-connection-string `
           --resource-group $ResourceGroupName `
           --name $storageAccountName `
           --query "connectionString" `
           --output tsv)

       # Create Office 365 API Connection first
		Write-Host "Creating Office 365 API Connection..." -ForegroundColor Yellow
		$office365ConnectionName = "$BotName-office365"
		$apiId = "/subscriptions/$SubscriptionId/providers/Microsoft.Web/locations/$Location/managedApis/office365"

		$connectionProperties = @{
			api = @{
				id = $apiId
			}
			displayName = "Office 365"
		} | ConvertTo-Json -Depth 10 | Set-Content "connection.json"

		# Create and capture the connection response
		$connection = az resource create `
			--resource-group $ResourceGroupName `
			--name $office365ConnectionName `
			--resource-type "Microsoft.Web/connections" `
			--properties "@connection.json" | ConvertFrom-Json

		Remove-Item "connection.json" -ErrorAction SilentlyContinue

		# Store the connection ID
		$connectionId = $connection.id
		Write-Host "Created connection with ID: $connectionId" -ForegroundColor Green

		# Check and wait for O365 connection authorization
		$connectionStatus = az resource show `
			--resource-group $ResourceGroupName `
			--name $office365ConnectionName `
			--resource-type "Microsoft.Web/connections" `
			--query "properties.statuses[0].status" -o tsv

		if ($connectionStatus -ne "Connected") {
			Write-Host "`nIMPORTANT: The Office 365 API Connection needs to be authorized." -ForegroundColor Yellow
			Write-Host "Please complete these steps in the Azure Portal:" -ForegroundColor Yellow
			Write-Host "1. Go to Resource Group: $ResourceGroupName" -ForegroundColor Yellow
			Write-Host "2. Find and click the Office 365 API Connection named: $office365ConnectionName" -ForegroundColor Yellow
			Write-Host "3. Click 'Edit API Connection'" -ForegroundColor Yellow
			Write-Host "4. Click 'Authorize' and sign in with your Office 365 account" -ForegroundColor Yellow
			Write-Host "5. Click 'Save'" -ForegroundColor Yellow
			
			Write-Host "`nPress Enter after you've completed the authorization in the Azure Portal..." -ForegroundColor Cyan
			$null = Read-Host

			# Verify the connection after authorization
			$connectionStatus = az resource show `
				--resource-group $ResourceGroupName `
				--name $office365ConnectionName `
				--resource-type "Microsoft.Web/connections" `
				--query "properties.statuses[0].status" -o tsv

			if ($connectionStatus -ne "Connected") {
				Write-Error "Connection still not authorized. Please verify the authorization was successful."
				exit 1
			}
			
			Write-Host "Office 365 connection successfully authorized!" -ForegroundColor Green
		}

		# Now create Logic App with authorized connection
		Write-Host "Creating Logic App..." -ForegroundColor Yellow
		$logicAppName = "$BotName-logic"

		$initialWorkflow = @{
			definition = @{
				'$schema' = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowDefinition.json#"
				contentVersion = "1.0.0.0"
				parameters = @{
					'$connections' = @{
						defaultValue = @{
							office365 = @{
								connectionId = $connectionId
								connectionName = $office365ConnectionName
								id = $apiId
							}
						}
						type = "Object"
					}
				}
				triggers = @{}
				actions = @{}
				outputs = @{}
			}
		} | ConvertTo-Json -Depth 10 | Set-Content "initial-workflow.json"

		az logic workflow create `
			--name $logicAppName `
			--resource-group $ResourceGroupName `
			--location $Location `
			--definition "@initial-workflow.json"

		Remove-Item "initial-workflow.json" -ErrorAction SilentlyContinue

       # Check for existing bot channels registration
       Write-Host "Checking for existing bot channels registration..." -ForegroundColor Yellow
       $existingBot = az bot show --name $BotName --resource-group $ResourceGroupName 2>$null
       
       if ($existingBot) {
           Write-Host "Updating existing bot channels registration..." -ForegroundColor Yellow
           $botChannels = az bot update `
               --name $BotName `
               --resource-group $ResourceGroupName `
               --appid $appId `
               --endpoint "https://$BotName.azurewebsites.net/api/messages" | ConvertFrom-Json
       } else {
           Write-Host "Creating new bot channels registration..." -ForegroundColor Yellow
           $botChannels = az bot create `
               --resource-group $ResourceGroupName `
               --name $BotName `
               --appid $appId `
               --app-type "MultiTenant" `
               --endpoint "https://$BotName.azurewebsites.net/api/messages" `
               --sku "F0" `
               --output json | ConvertFrom-Json
       }

       # Check for existing App Service Plan
       $planName = "$BotName-plan"
       Write-Host "Checking for existing App Service Plan..." -ForegroundColor Yellow
       $existingPlan = az appservice plan show --name $planName --resource-group $ResourceGroupName 2>$null

       if (-not $existingPlan) {
           Write-Host "Creating new App Service Plan..." -ForegroundColor Yellow
           az appservice plan create `
               --name $planName `
               --resource-group $ResourceGroupName `
               --location $Location `
               --sku "F1"
       } else {
           Write-Host "Using existing App Service Plan: $planName" -ForegroundColor Yellow
       }

       # Check for existing Function App
       $functionAppName = "$($BotName.ToLower() -replace '[^a-z0-9-]', '')-func"
       Write-Host "Checking for existing Function App..." -ForegroundColor Yellow
       $existingFunctionApp = az functionapp show --name $functionAppName --resource-group $ResourceGroupName 2>$null

       if (-not $existingFunctionApp) {
           Write-Host "Creating new Function App..." -ForegroundColor Yellow
           $functionApp = az functionapp create `
               --name $functionAppName `
               --resource-group $ResourceGroupName `
               --storage-account $storageAccountName `
               --runtime "node" `
               --runtime-version "20" `
               --functions-version 4 `
               --consumption-plan-location $Location `
               --os-type Linux `
               --output json | ConvertFrom-Json
       } else {
           Write-Host "Using existing Function App: $functionAppName" -ForegroundColor Yellow
           $functionApp = $existingFunctionApp | ConvertFrom-Json
       }

       # Get function URL and key
       Write-Host "Getting function URL and key..." -ForegroundColor Yellow
       $functionUrl = "https://$functionAppName.azurewebsites.net/api/handleapproval"
       $functionKey = (az functionapp keys list -g $ResourceGroupName -n $functionAppName --query "functionKeys.default" -o tsv)

       # Set Function App configuration
       Write-Host "Configuring Function App settings..." -ForegroundColor Yellow
       $functionAppSettings = @{
           "MicrosoftAppId" = $appId
           "MicrosoftAppPassword" = $secret.password
           "MicrosoftAppTenantId" = $TenantId
           "AzureStorageConnectionString" = $storageConnString
           "NODE_ENV" = "production"
       }

       # Apply settings to Function App
       foreach ($setting in $functionAppSettings.GetEnumerator()) {
           az functionapp config appsettings set `
               --name $functionAppName `
               --resource-group $ResourceGroupName `
               --settings "$($setting.Key)=$($setting.Value)"
       }

       # Check for existing Web App
       $webAppName = $BotName
       Write-Host "Checking for existing Web App..." -ForegroundColor Yellow
       $existingWebApp = az webapp show --name $webAppName --resource-group $ResourceGroupName 2>$null

       if (-not $existingWebApp) {
           Write-Host "Creating new Web App..." -ForegroundColor Yellow
           az webapp create `
               --resource-group $ResourceGroupName `
               --plan $planName `
               --name $webAppName `
               --runtime "NODE:20LTS"
       } else {
           Write-Host "Using existing Web App: $webAppName" -ForegroundColor Yellow
       }

       # Configure CORS settings for the web app
       Write-Host "Configuring Web App CORS settings..." -ForegroundColor Yellow
       az webapp cors add `
           --resource-group $ResourceGroupName `
           --name $BotName `
           --allowed-origins "https://teams.microsoft.com" `
           --allowed-origins "https://*.teams.microsoft.com" `
           --allowed-origins "https://outlook.office.com" `
           --allowed-origins "https://outlook.office365.com"

       # Set Web App configuration with critical environment variables
       Write-Host "Configuring Web App settings..." -ForegroundColor Yellow
       $webAppSettings = @{
           "MicrosoftAppId" = $appId
           "MicrosoftAppPassword" = $secret.password
           "MicrosoftAppTenantId" = $TenantId
           "AzureStorageConnectionString" = $storageConnString
           "NODE_ENV" = "production"
           "FUNCTIONAPP_URL" = $functionUrl
           "FUNCTIONAPP_KEY" = $functionKey
       }

       # Apply settings to Web App
       foreach ($setting in $webAppSettings.GetEnumerator()) {
           az webapp config appsettings set `
               --resource-group $ResourceGroupName `
               --name $webAppName `
               --settings "$($setting.Key)=$($setting.Value)"
       }

       return @{
           BotId = $appId
           BotPassword = $secret.password
           StorageConnString = $storageConnString
           WebAppName = $webAppName
           FunctionAppName = $functionApp.name
           StorageAccountName = $storageAccountName
           AppServicePlanName = $planName
           LogicAppName = $logicAppName
           Office365ConnectionName = $office365ConnectionName
		   Office365ConnectionId = $connectionId
		   Office365ApiId = $apiId
       }
   }
   catch {
       Write-Error "Error creating Azure resources: $_"
       throw
   }
}

function Clone-Repository {
    param(
        [string]$RepoUrl,
        [string]$LocalPath
    )
    
    Write-Host "`n=== Cloning Repository ===" -ForegroundColor Cyan
    
    # Remove existing directory if it exists
    if (Test-Path $LocalPath) {
        Write-Host "Removing existing repository directory..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $LocalPath
    }
    
    # Create parent directory if it doesn't exist
    $parentDir = Split-Path -Parent $LocalPath
    if (-not (Test-Path $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force
    }
    
    # Clone the repository
    Write-Host "Cloning repository from $RepoUrl to $LocalPath..." -ForegroundColor Yellow
    git clone $RepoUrl $LocalPath
    
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to clone repository from $RepoUrl"
    }
    
    Write-Host "Repository cloned successfully!" -ForegroundColor Green
    
    # Install npm dependencies for both bot and function
    Write-Host "Installing npm dependencies for bot..." -ForegroundColor Yellow
    Set-Location $LocalPath
    npm install
    
    Write-Host "Installing npm dependencies for function..." -ForegroundColor Yellow
    Set-Location (Join-Path $LocalPath "functions")
    npm install
    
    Set-Location $LocalPath
}

function Deploy-CodeToAzure {
    param(
        [string]$RepoPath,
        [string]$WebAppName,
        [string]$FunctionAppName,
        [string]$ResourceGroupName,
        [string]$SubscriptionId,
        [string]$Location,
        [string]$Office365ConnectionName,
        [string]$Office365ConnectionId,
        [string]$Office365ApiId
    )
    
    Write-Host "`n=== Deploying Code to Azure Resources ===" -ForegroundColor Cyan
    
    # Function App Deployment
    Write-Host "Preparing Function App deployment..." -ForegroundColor Yellow
    
    # Create temp directory structure for function app
    $functionTempPath = Join-Path $RepoPath "functions\temp_deploy"
    Remove-Item -Recurse -Force $functionTempPath -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path "$functionTempPath\handleapproval" -Force
    
    # Copy function files
    Write-Host "Copying function files..." -ForegroundColor Yellow
    Copy-Item "$RepoPath\functions\handleapproval\function.json" "$functionTempPath\handleapproval\"
    Copy-Item "$RepoPath\functions\handleapproval\handleapproval.js" "$functionTempPath\handleapproval\"
    Copy-Item "$RepoPath\functions\handleapproval\index.js" "$functionTempPath\handleapproval\"
    Copy-Item "$RepoPath\functions\package.json" "$functionTempPath\"
    Copy-Item "$RepoPath\functions\package-lock.json" "$functionTempPath\"
    Copy-Item -Recurse "$RepoPath\functions\node_modules" "$functionTempPath\"
    
    # Create zip file for function app
    Write-Host "Creating function app zip package..." -ForegroundColor Yellow
    Set-Location $functionTempPath
    $functionZipPath = Join-Path $RepoPath "functions\functionapp.zip"
    Remove-Item -Force $functionZipPath -ErrorAction SilentlyContinue
    Compress-Archive -Path "*" -DestinationPath $functionZipPath -Force
    Set-Location $RepoPath
    
    # Deploy function app
    Write-Host "Deploying Function App..." -ForegroundColor Yellow
    az functionapp deployment source config-zip `
        -g $ResourceGroupName `
        -n $FunctionAppName `
        --src $functionZipPath
    
    # Get function key
    Write-Host "Retrieving Function App key..." -ForegroundColor Yellow
    $functionKey = az functionapp keys list `
        -g $ResourceGroupName `
        -n $FunctionAppName `
        --query "functionKeys.default" `
        -o tsv
    
    Write-Host "Function Key: $functionKey" -ForegroundColor Green
    
    # Bot Web App Deployment
    Write-Host "Preparing Bot Web App deployment..." -ForegroundColor Yellow
    
    # Create temp directory for bot deployment
    $botTempPath = Join-Path $RepoPath "temp_deploy_bot"
    Remove-Item -Recurse -Force $botTempPath -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $botTempPath -Force
    
    # Copy bot files
    Write-Host "Copying bot files..." -ForegroundColor Yellow
    Copy-Item "$RepoPath\bot.js" $botTempPath
    Copy-Item "$RepoPath\index.js" $botTempPath
    Copy-Item "$RepoPath\package.json" $botTempPath
    Copy-Item "$RepoPath\package-lock.json" $botTempPath
    Copy-Item "$RepoPath\web.config" $botTempPath -ErrorAction SilentlyContinue
    Copy-Item -Recurse "$RepoPath\node_modules" $botTempPath
    
    # Create zip file for bot web app
    Write-Host "Creating bot web app zip package..." -ForegroundColor Yellow
    Set-Location $botTempPath
    $botZipPath = Join-Path $RepoPath "botapp.zip"
    Remove-Item -Force $botZipPath -ErrorAction SilentlyContinue
    Compress-Archive -Path "*" -DestinationPath $botZipPath -Force
    Set-Location $RepoPath
    
    # Deploy bot web app
    Write-Host "Deploying Bot Web App..." -ForegroundColor Yellow
    az webapp deployment source config-zip `
        --resource-group $ResourceGroupName `
        --name $WebAppName `
        --src $botZipPath

    Write-Host "Connection Details being used:" -ForegroundColor Yellow
    Write-Host "Connection Name: $office365ConnectionName"
    Write-Host "Connection ID: $connectionId"
    Write-Host "API ID: $apiId"

    # Update Logic App Workflow with actual workflow
    Write-Host "Updating Logic App workflow..." -ForegroundColor Yellow
    $logicAppName = "$WebAppName-logic"
    $workflowFile = Join-Path $RepoPath "logic-app-workflow.json"

    if (Test-Path $workflowFile) {
        $workflow = Get-Content $workflowFile | ConvertFrom-Json
        
        # Verify/Show current values
        Write-Host "Using connection name: $Office365ConnectionName" -ForegroundColor Yellow
        
        # Update Office 365 connection references
        Write-Host "Connection Details being used:" -ForegroundColor Yellow
        Write-Host "Connection Name: $Office365ConnectionName"
        Write-Host "Connection ID: $Office365ConnectionId"
        Write-Host "API ID: $Office365ApiId"

        # Update connection parameters
        $workflow.definition.parameters.'$connections' = @{
            defaultValue = @{
                office365 = @{
                    connectionId = $Office365ConnectionId
                    connectionName = $Office365ConnectionName
                    id = $Office365ApiId
                }
            }
            type = "Object"
        }

        # Add BotName parameter
        $workflow.definition.parameters.BotName = @{
            type = "String"
            defaultValue = $WebAppName
        }
        
        # Save and deploy updated workflow
        $workflow | ConvertTo-Json -Depth 30 | Set-Content "workflow-deploy.json"
        
        Write-Host "Deploying Logic App with updated workflow..." -ForegroundColor Yellow
        az logic workflow create `
            --resource-group $ResourceGroupName `
            --name $logicAppName `
            --definition "@workflow-deploy.json" `
            --location $Location
            
        Remove-Item "workflow-deploy.json" -ErrorAction SilentlyContinue
    }
    
    # Update bot web app settings with function key
    Write-Host "Updating Bot Web App settings with Function Key..." -ForegroundColor Yellow
    az webapp config appsettings set `
        --resource-group $ResourceGroupName `
        --name $WebAppName `
        --settings FUNCTIONAPP_KEY=$functionKey
    
    # Cleanup
    Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $functionTempPath -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $botTempPath -ErrorAction SilentlyContinue
    Remove-Item -Force $functionZipPath -ErrorAction SilentlyContinue
    Remove-Item -Force $botZipPath -ErrorAction SilentlyContinue
    
    Write-Host "Code deployment completed successfully!" -ForegroundColor Green
}

function New-TeamsManifest {
    param(
        [string]$RepoPath,
        [string]$BotName,
        [string]$BotId
    )
    
    Write-Host "`n=== Creating Teams App Manifest ===" -ForegroundColor Cyan
    
    $WebsiteUrl = "https://$BotName.azurewebsites.net"
    
    # Create manifest directory
    $manifestDir = Join-Path $RepoPath "teams-manifest"
    New-Item -ItemType Directory -Path $manifestDir -Force | Out-Null
    
    # Create manifest content
    $manifest = @{
        "`$schema" = "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json"
        manifestVersion = "1.14"
        version = "1.0.0"
        id = $BotId
        packageName = "com.microsoft.teams.prabot"
        developer = @{
            name = "PRA Approval Bot Team"
            websiteUrl = $WebsiteUrl
            privacyUrl = "$WebsiteUrl/privacy"
            termsOfUseUrl = "$WebsiteUrl/termsofuse"
        }
        name = @{
            short = $BotName
            full = "$BotName - PRA Access Management"
        }
        description = @{
            short = "Manages PRA access requests"
            full = "This bot helps manage and process approval requests for BeyondTrust Privileged Remote Access (PRA) in a secure and efficient manner."
        }
        icons = @{
            color = "color.png"
            outline = "outline.png"
        }
        accentColor = "#FFFFFF"
        bots = @(
            @{
                botId = $BotId
                scopes = @(
                    "team",
                    "personal"
                )
                supportsFiles = $false
                isNotificationOnly = $false
            }
        )
        permissions = @(
            "messageTeamMembers"
        )
        validDomains = @(
            "$BotName.azurewebsites.net"
        )
    }
    
    # Save manifest and validate JSON
    $manifestPath = Join-Path $manifestDir "manifest.json"
    $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8
    
    # Copy icons
    Copy-Item -Path (Join-Path $RepoPath "color.png") -Destination $manifestDir
    Copy-Item -Path (Join-Path $RepoPath "outline.png") -Destination $manifestDir
    
    # Create zip
    $zipPath = Join-Path $RepoPath "$BotName-teams-package.zip"
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }
    
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($manifestDir, $zipPath)
    
    Write-Host "Teams manifest package created at: $zipPath" -ForegroundColor Green
    
    # Cleanup manifest directory
    Remove-Item -Path $manifestDir -Recurse -Force
    
    return $zipPath
}

# Main script execution starts here
try {
    # Install dependencies
    Install-RequiredModules

    # Connect to Azure
    Connect-ToAzure

    # Create Azure resources
    $resources = New-AzureResources

    # Set environment variables for Function App
    $envVars = @{
        "BEYONDTRUST_BASE_URL" = $BeyondTrustBaseUrl
        "BEYONDTRUST_CLIENT_ID" = $BeyondTrustClientId
        "BEYONDTRUST_CLIENT_SECRET" = $BeyondTrustClientSecret
        "MicrosoftAppId" = $resources.BotId
        "MicrosoftAppPassword" = $resources.BotPassword
        "AzureStorageConnectionString" = $resources.StorageConnString
    }

    foreach ($var in $envVars.GetEnumerator()) {
        az functionapp config appsettings set `
            --name "$($BotName)-func" `
            --resource-group $ResourceGroupName `
            --settings "$($var.Key)=$($var.Value)"
    }

    # Clone the repository
    Clone-Repository -RepoUrl $GitRepoUrl -LocalPath $LocalRepoPath

    Deploy-CodeToAzure `
    -RepoPath $LocalRepoPath `
    -WebAppName $resources.WebAppName `
    -FunctionAppName $resources.FunctionAppName `
    -ResourceGroupName $ResourceGroupName `
    -SubscriptionId $SubscriptionId `
    -Location $Location `
    -Office365ConnectionName $resources.Office365ConnectionName `
    -Office365ConnectionId $resources.Office365ConnectionId `
    -Office365ApiId $resources.Office365ApiId
	
	# Create Teams manifest package
    Write-Host "`n=== Creating Teams Manifest Package ===" -ForegroundColor Cyan
    $manifestPath = New-TeamsManifest `
        -RepoPath $LocalRepoPath `
        -BotName $BotName `
        -BotId $resources.BotId

    Write-Host "`n=== Deployment Summary ===" -ForegroundColor Cyan
    Write-Host "Bot Deployed Successfully!" -ForegroundColor Green
    Write-Host "Teams manifest package created at: $manifestPath" -ForegroundColor Green
    
    Write-Host "`nNext Steps:" -ForegroundColor Yellow
    Write-Host "1. Create a new team in Microsoft Teams manually" -ForegroundColor Yellow
    Write-Host "2. Go to Teams Admin Center" -ForegroundColor Yellow
    Write-Host "3. Upload the manifest package from: $manifestPath" -ForegroundColor Yellow
    Write-Host "4. Add the bot to your team" -ForegroundColor Yellow

    Write-Host "`nDeployment completed successfully!" -ForegroundColor Green
}	
catch {
    Write-Error "An error occurred during deployment: $_"
    exit 1
}
