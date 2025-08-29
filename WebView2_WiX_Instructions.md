# WiX MSI Package Instructions for WebView2 Integration

There are two ways to integrate WebView2 components into your WiX-based MSI package for YoutubePlus:

## Option 1: Automatic Integration (Recommended)

A PowerShell script has been created that will automatically apply all the necessary changes to your WiX project:

1. Open PowerShell and navigate to your project directory:
   ```powershell
   cd C:\Users\richi\source\repos\YoutubePlus
   ```

2. Run the automatic integration script:
   ```powershell
   .\ApplyWebView2WiXChanges.ps1
   ```

3. The script will:
   - Copy WebView2Loader.dll to your project if needed
   - Download MicrosoftEdgeWebview2Setup.exe if missing
   - Add WebView2Components.wxs to your WiX project
   - Update Package.wxs with all necessary WebView2 configuration
   - Set proper permissions in your MSI package

4. After the script completes, build your WiX project to create the MSI

## Option 2: Manual Integration

If you prefer to make the changes manually, follow these steps:

### 1. Include WebView2Components.wxs in Your Project

The `WebView2Components.wxs` file has been created for you. Add this file to your WiX project:

1. Right-click on your WiX project in Solution Explorer
2. Select "Add" → "Existing Item"
3. Browse to and select `WebView2Components.wxs`

### 2. Reference the Component Group in Your Main Package.wxs

Open your main `Package.wxs` file and make the following changes:

#### Add the ComponentGroup Reference

Find the `<Feature>` element in your Package.wxs and add the WebView2Components ComponentGroup:

```xml
<Feature Id="ProductFeature" Title="YoutubePlus" Level="1">
    <ComponentGroupRef Id="ProductComponents" />
    <!-- Add this line -->
    <ComponentGroupRef Id="WebView2Components" />
</Feature>
```

#### Add WebView2 Installation Verification

Add the following inside your `<Product>` element (but outside any Feature elements):

```xml
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
```

### 3. Ensure Proper Permissions

To ensure your MSI can create directories in the AppData folder, add the following to your Package.wxs file:

```xml
<!-- Request elevated privileges -->
<Package InstallerVersion="200" Compressed="yes" InstallScope="perMachine" 
         InstallPrivileges="elevated" />
```

### 4. Verify Setup and Build

1. Run the `VerifyWebView2Setup.ps1` script to check all requirements
2. Build your WiX project to create the MSI
3. Test the MSI on a clean machine to verify WebView2 is properly installed

## Troubleshooting

If you encounter issues with WebView2 not being properly included or detected:

1. Verify both WebView2Loader.dll and MicrosoftEdgeWebview2Setup.exe are in your project root
2. Check that the file paths in WebView2Components.wxs are correct
3. Ensure your WiX project includes WebView2Components.wxs in the build
4. Review the custom action implementation for any syntax errors

### Common Issues and Solutions

1. **"WebView2 not initialized" error after installation**
   - Verify WebView2Loader.dll is properly included in the MSI and installed to the application directory
   - Ensure the application has permission to create the WebView2 user data folder in AppData

2. **WebView2 installation fails during MSI installation**
   - Check if UAC (User Account Control) is blocking the installer
   - Try running the MSI as administrator
   - Verify internet connectivity for downloading WebView2 components if needed

3. **WiX build errors**
   - Make sure all file paths in WebView2Components.wxs are valid
   - Ensure there are no duplicate component IDs in your WiX files

For additional assistance, consult the [WebView2 documentation](https://docs.microsoft.com/en-us/microsoft-edge/webview2/) or [WiX documentation](https://wixtoolset.org/documentation/).
