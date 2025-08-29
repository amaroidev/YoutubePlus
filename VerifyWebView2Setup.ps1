# VerifyWebView2Setup.ps1
# Script to verify and prepare WebView2 requirements for WiX MSI packaging
# For YoutubePlus application

Write-Host "=== WebView2 Setup Verification for WiX MSI Package ===" -ForegroundColor Cyan

# Paths
$projectRoot = Split-Path -Parent $PSCommandPath
$wixProjectPath = Join-Path $projectRoot "YoutubePlus.Setup"
$wxsFilePath = Join-Path $wixProjectPath "Package.wxs"
$webViewSetupPath = Join-Path $projectRoot "MicrosoftEdgeWebview2Setup.exe"
$webViewLoaderPath = Join-Path $projectRoot "WebView2Loader.dll"

# Step 1: Verify WebView2Loader.dll exists
Write-Host "`nChecking for WebView2Loader.dll..." -ForegroundColor Yellow
if (Test-Path $webViewLoaderPath) {
    Write-Host "WebView2Loader.dll found at $webViewLoaderPath" -ForegroundColor Green
} else {
    Write-Host "WebView2Loader.dll not found!" -ForegroundColor Red
    
    # Check in bin/Release or x64/Release folders
    $possiblePaths = @(
        (Join-Path $projectRoot "Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "x64\Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "bin\Release\WebView2Loader.dll")
    )
    
    $found = $false
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            Write-Host "Found WebView2Loader.dll at $path" -ForegroundColor Green
            Write-Host "Copying to project root..." -ForegroundColor Yellow
            Copy-Item $path $projectRoot
            $found = $true
            break
        }
    }
    
    if (-not $found) {
        Write-Host "ERROR: WebView2Loader.dll not found in any expected location." -ForegroundColor Red
        Write-Host "Please make sure the WebView2 SDK is properly installed and the application builds successfully." -ForegroundColor Red
    }
}

# Step 2: Verify MicrosoftEdgeWebview2Setup.exe exists
Write-Host "`nChecking for MicrosoftEdgeWebview2Setup.exe..." -ForegroundColor Yellow
if (Test-Path $webViewSetupPath) {
    Write-Host "MicrosoftEdgeWebview2Setup.exe found at $webViewSetupPath" -ForegroundColor Green
} else {
    Write-Host "MicrosoftEdgeWebview2Setup.exe not found!" -ForegroundColor Red
    Write-Host "Downloading from Microsoft..." -ForegroundColor Yellow
    
    try {
        $url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
        Invoke-WebRequest -Uri $url -OutFile $webViewSetupPath
        Write-Host "Downloaded successfully to $webViewSetupPath" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to download WebView2 installer." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

# Step 3: Check WiX project file for WebView2 components
Write-Host "`nChecking WiX Package.wxs file for WebView2 components..." -ForegroundColor Yellow
if (Test-Path $wxsFilePath) {
    $wxsContent = Get-Content $wxsFilePath -Raw
    
    # Check for WebView2Loader.dll
    if ($wxsContent -match "WebView2Loader.dll") {
        Write-Host "WebView2Loader.dll is included in the WiX package file" -ForegroundColor Green
    } else {
        Write-Host "WARNING: WebView2Loader.dll might not be included in your WiX package file" -ForegroundColor Red
        Write-Host "Please ensure you add the following to your WiX file (inside a ComponentGroup):" -ForegroundColor Yellow
        Write-Host @"
<Component Id="WebView2Loader.dll" Guid="*">
    <File Id="WebView2Loader.dll" Source="`$(var.ProjectDir)\..\WebView2Loader.dll" KeyPath="yes" />
</Component>
"@ -ForegroundColor White
    }
    
    # Check for MicrosoftEdgeWebview2Setup.exe
    if ($wxsContent -match "MicrosoftEdgeWebview2Setup.exe") {
        Write-Host "MicrosoftEdgeWebview2Setup.exe is included in the WiX package file" -ForegroundColor Green
    } else {
        Write-Host "WARNING: MicrosoftEdgeWebview2Setup.exe might not be included in your WiX package file" -ForegroundColor Red
        Write-Host "Please ensure you add the following to your WiX file (inside a ComponentGroup):" -ForegroundColor Yellow
        Write-Host @"
<Component Id="MicrosoftEdgeWebview2Setup.exe" Guid="*">
    <File Id="MicrosoftEdgeWebview2Setup.exe" Source="`$(var.ProjectDir)\..\MicrosoftEdgeWebview2Setup.exe" KeyPath="yes" />
</Component>
"@ -ForegroundColor White
    }
    
    # Check for WebView2 installation custom action
    if ($wxsContent -match "CheckWebView2" -or $wxsContent -match "InstallWebView2") {
        Write-Host "WebView2 installation verification custom action found in WiX file" -ForegroundColor Green
    } else {
        Write-Host "WARNING: No WebView2 installation verification custom action found in WiX file" -ForegroundColor Red
        Write-Host "Consider adding a custom action to verify WebView2 is installed:" -ForegroundColor Yellow
        Write-Host @"
<!-- Add to your Product element -->
<Property Id="WEBVIEW2INSTALLED">
    <RegistrySearch Id="WebView2InstalledSearch" Root="HKLM" Key="SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" Name="pv" Type="raw" />
</Property>

<CustomAction Id="InstallWebView2" FileKey="MicrosoftEdgeWebview2Setup.exe" ExeCommand="/silent /install" Return="check" />

<InstallExecuteSequence>
    <Custom Action="InstallWebView2" After="InstallFiles">NOT WEBVIEW2INSTALLED</Custom>
</InstallExecuteSequence>
"@ -ForegroundColor White
    }
} else {
    Write-Host "ERROR: Could not find WiX Package.wxs file at $wxsFilePath" -ForegroundColor Red
}

# Step 4: Generate WiX Component for WebView2 if needed
Write-Host "`nWould you like to generate WiX component snippets for WebView2 files? (Y/N)" -ForegroundColor Cyan
$response = Read-Host
if ($response -eq "Y" -or $response -eq "y") {
    $snippetsPath = Join-Path $projectRoot "WebView2WiXSnippets.txt"
    
    @"
<!-- Add these components to your ComponentGroup in the WiX file -->

<!-- WebView2Loader.dll Component -->
<Component Id="WebView2Loader.dll" Guid="*">
    <File Id="WebView2Loader.dll" Source="`$(var.ProjectDir)\..\WebView2Loader.dll" KeyPath="yes" />
</Component>

<!-- MicrosoftEdgeWebview2Setup.exe Component -->
<Component Id="MicrosoftEdgeWebview2Setup.exe" Guid="*">
    <File Id="MicrosoftEdgeWebview2Setup.exe" Source="`$(var.ProjectDir)\..\MicrosoftEdgeWebview2Setup.exe" KeyPath="yes" />
</Component>

<!-- WebView2 Installation Check and Install Custom Action -->
<!-- Add to your Product element -->
<Property Id="WEBVIEW2INSTALLED">
    <RegistrySearch Id="WebView2InstalledSearch" Root="HKLM" Key="SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" Name="pv" Type="raw" />
</Property>

<CustomAction Id="InstallWebView2" FileKey="MicrosoftEdgeWebview2Setup.exe" ExeCommand="/silent /install" Return="check" />

<InstallExecuteSequence>
    <Custom Action="InstallWebView2" After="InstallFiles">NOT WEBVIEW2INSTALLED</Custom>
</InstallExecuteSequence>
"@ | Out-File -FilePath $snippetsPath

    Write-Host "WiX component snippets for WebView2 files generated at:" -ForegroundColor Green
    Write-Host $snippetsPath -ForegroundColor Green
}

Write-Host "`n=== WebView2 Setup Verification Complete ===" -ForegroundColor Cyan
Write-Host "Make sure to include these files in your MSI and grant appropriate permissions in your installer." -ForegroundColor Yellow
