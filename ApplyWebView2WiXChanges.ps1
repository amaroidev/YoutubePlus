# ApplyWebView2WiXChanges.ps1
# Script to automatically apply WebView2 integration to WiX project
# For YoutubePlus application

Write-Host "=== Automatically Applying WebView2 Integration to WiX MSI Package ===" -ForegroundColor Cyan

# Paths
$projectRoot = Split-Path -Parent $PSCommandPath
$wixProjectPath = Join-Path $projectRoot "YoutubePlus.Setup"
$wxsFilePath = Join-Path $wixProjectPath "Package.wxs"
$webViewComponentsPath = Join-Path $wixProjectPath "WebView2Components.wxs"
$webViewSetupPath = Join-Path $projectRoot "MicrosoftEdgeWebview2Setup.exe"
$webViewLoaderPath = Join-Path $projectRoot "WebView2Loader.dll"
$wixProjPath = Join-Path $wixProjectPath "YoutubePlus.Setup.wixproj"

# Step 1: Ensure WebView2Loader.dll exists
Write-Host "`nChecking for WebView2Loader.dll..." -ForegroundColor Yellow
if (-not (Test-Path $webViewLoaderPath)) {
    # Check in bin/Release or x64/Release folders
    $possiblePaths = @(
        (Join-Path $projectRoot "Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "x64\Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "bin\Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "YoutubePlus\Release\WebView2Loader.dll"),
        (Join-Path $projectRoot "YoutubePlus\x64\Release\WebView2Loader.dll")
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
        Write-Host "Please build your project first to generate WebView2Loader.dll" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "WebView2Loader.dll found at $webViewLoaderPath" -ForegroundColor Green
}

# Step 2: Ensure MicrosoftEdgeWebview2Setup.exe exists
Write-Host "`nChecking for MicrosoftEdgeWebview2Setup.exe..." -ForegroundColor Yellow
if (-not (Test-Path $webViewSetupPath)) {
    Write-Host "MicrosoftEdgeWebview2Setup.exe not found. Downloading from Microsoft..." -ForegroundColor Yellow
    
    try {
        $url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
        Invoke-WebRequest -Uri $url -OutFile $webViewSetupPath
        Write-Host "Downloaded successfully to $webViewSetupPath" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to download WebView2 installer." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "MicrosoftEdgeWebview2Setup.exe found at $webViewSetupPath" -ForegroundColor Green
}

# Step 3: Add WebView2Components.wxs to the WiX project if it's not already included
Write-Host "`nAdding WebView2Components.wxs to WiX project..." -ForegroundColor Yellow
if (Test-Path $wixProjPath) {
    # Read the project file
    [xml]$wixProj = Get-Content $wixProjPath
    
    # Check if WebView2Components.wxs is already included
    $isIncluded = $false
    foreach ($compile in $wixProj.Project.ItemGroup.Compile) {
        if ($compile.Include -match "WebView2Components.wxs") {
            $isIncluded = $true
            break
        }
    }
    
    # Add WebView2Components.wxs if not already included
    if (-not $isIncluded) {
        Write-Host "Adding WebView2Components.wxs to the WiX project file..." -ForegroundColor Yellow
        
        # Create a new ItemGroup if there isn't one with Compile items
        $itemGroup = $null
        foreach ($group in $wixProj.Project.ItemGroup) {
            if ($group.Compile) {
                $itemGroup = $group
                break
            }
        }
        
        if ($itemGroup -eq $null) {
            $itemGroup = $wixProj.CreateElement("ItemGroup", $wixProj.Project.NamespaceURI)
            [void]$wixProj.Project.AppendChild($itemGroup)
        }
        
        # Create and add the Compile element
        $compile = $wixProj.CreateElement("Compile", $wixProj.Project.NamespaceURI)
        $compile.SetAttribute("Include", "WebView2Components.wxs")
        [void]$itemGroup.AppendChild($compile)
        
        # Save the changes
        $wixProj.Save($wixProjPath)
        Write-Host "WebView2Components.wxs added to WiX project file successfully" -ForegroundColor Green
    } else {
        Write-Host "WebView2Components.wxs is already included in the WiX project file" -ForegroundColor Green
    }
} else {
    Write-Host "WARNING: Could not find WiX project file at $wixProjPath" -ForegroundColor Red
    Write-Host "You will need to manually add WebView2Components.wxs to your WiX project" -ForegroundColor Yellow
}

# Step 4: Update Package.wxs
Write-Host "`nUpdating Package.wxs with WebView2 components..." -ForegroundColor Yellow
if (Test-Path $wxsFilePath) {
    # Read the current content
    $wxsContent = Get-Content $wxsFilePath -Raw
    
    # Check if already modified
    $modified = $false
    if ($wxsContent -match "WebView2Components") {
        $modified = $true
        Write-Host "Package.wxs already contains WebView2Components references" -ForegroundColor Green
    }
    
    if (-not $modified) {
        # Backup the original file
        $backupPath = "$wxsFilePath.backup"
        Copy-Item $wxsFilePath $backupPath
        Write-Host "Created backup of Package.wxs at $backupPath" -ForegroundColor Green
        
        # Add ComponentGroupRef to Feature element
        if ($wxsContent -match "<Feature.+?ProductFeature.+?>") {
            $featureMatch = $Matches[0]
            $featureBlock = $wxsContent -split $featureMatch, 2
            
            if ($featureBlock.Count -gt 1) {
                $closingFeatureIndex = $featureBlock[1].IndexOf("</Feature>")
                if ($closingFeatureIndex -ge 0) {
                    $featureContent = $featureBlock[1].Substring(0, $closingFeatureIndex)
                    $componentGroupRef = "`n    <ComponentGroupRef Id=`"WebView2Components`" />"
                    
                    # Check if the ComponentGroupRef is already there
                    if ($featureContent -notmatch "WebView2Components") {
                        $newFeatureContent = $featureContent + $componentGroupRef
                        $wxsContent = $featureBlock[0] + $featureMatch + $newFeatureContent + "</Feature>" + $featureBlock[1].Substring($closingFeatureIndex + 10)
                        
                        Write-Host "Added WebView2Components ComponentGroupRef to Feature element" -ForegroundColor Green
                    }
                }
            }
        }
        
        # Add WebView2 installation verification to Product element
        if ($wxsContent -match "<Product.+?>") {
            $webViewInstallationBlock = @"

    <!-- WebView2 Installation Check -->
    <Property Id="WEBVIEW2INSTALLED">
        <RegistrySearch Id="WebView2InstalledSearch" Root="HKLM" 
                        Key="SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" 
                        Name="pv" Type="raw" />
    </Property>

    <!-- Custom Action to install WebView2 if not present -->
    <CustomAction Id="InstallWebView2" 
                  FileKey="MicrosoftEdgeWebview2Setup.exe" 
                  ExeCommand="/silent /install" 
                  Return="check" />

    <!-- Execute WebView2 installation if needed -->
    <InstallExecuteSequence>
        <Custom Action="InstallWebView2" After="InstallFiles">NOT WEBVIEW2INSTALLED</Custom>
    </InstallExecuteSequence>
"@
            # Check if the installation block is already there
            if ($wxsContent -notmatch "WEBVIEW2INSTALLED") {
                # Find a good insertion point - after Product opening but before Feature
                if ($wxsContent -match "<Feature.+?ProductFeature.+?>") {
                    $featureMatch = $Matches[0]
                    $wxsContent = $wxsContent -replace $featureMatch, "$webViewInstallationBlock`n    $featureMatch"
                    
                    Write-Host "Added WebView2 installation verification to Product element" -ForegroundColor Green
                }
            }
        }
        
        # Ensure Package has elevated privileges
        if ($wxsContent -match "<Package.+?>") {
            $packageMatch = $Matches[0]
            
            # Check if it already has InstallScope and InstallPrivileges
            if ($packageMatch -notmatch "InstallScope" -or $packageMatch -notmatch "InstallPrivileges") {
                $newPackage = $packageMatch -replace "/>", " InstallScope=`"perMachine`" InstallPrivileges=`"elevated`" />"
                $wxsContent = $wxsContent -replace [regex]::Escape($packageMatch), $newPackage
                
                Write-Host "Updated Package element with elevated privileges" -ForegroundColor Green
            }
        }
        
        # Save the modified content
        Set-Content $wxsFilePath $wxsContent
        Write-Host "Package.wxs updated successfully" -ForegroundColor Green
    }
} else {
    Write-Host "ERROR: Could not find Package.wxs at $wxsFilePath" -ForegroundColor Red
}

Write-Host "`n=== WebView2 Integration Applied Successfully ===" -ForegroundColor Cyan
Write-Host "Your WiX project has been updated to include WebView2 components." -ForegroundColor Green
Write-Host "Please build your project to create the MSI package." -ForegroundColor Yellow
