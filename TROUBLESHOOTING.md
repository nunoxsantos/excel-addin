# Excel Add-in Installation Troubleshooting Guide

## Common Issues and Solutions

### 1. Mac Excel Installation Issues

**Problem**: Add-in won't install on Mac Excel
**Solutions**:

#### Method 1: Use the Sideload Package
1. Download `excel-addin-sideload.zip` from this repository
2. Extract the zip file
3. In Excel, go to **Insert** > **Office Add-ins**
4. Click **"Upload My Add-in"**
5. Select the extracted `manifest.xml` file

#### Method 2: Direct Manifest Download
1. Go to: https://nunoxsantos.github.io/excel-addin/manifest.xml
2. Right-click and "Save As" to download
3. In Excel, go to **Insert** > **Office Add-ins**
4. Click **"Upload My Add-in"**
5. Select the downloaded `manifest.xml` file

#### Method 3: Clear Excel Cache (Mac)
1. Close Excel completely
2. Open Finder
3. Press `Cmd + Shift + G`
4. Navigate to: `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`
5. Delete all contents in the Caches folder
6. Restart Excel and try again

### 2. Excel Version Requirements
- **Minimum**: Excel 2016 or later on Mac
- **Recommended**: Excel 2019 or later, or Microsoft 365
- **Check version**: Excel menu > About Microsoft Excel

### 3. Network and Security Issues

#### Check if GitHub Pages is accessible:
1. Open browser and go to: https://nunoxsantos.github.io/excel-addin/
2. You should see the add-in page load
3. If it doesn't load, there might be a network issue

#### Corporate Network Issues:
- Some corporate networks block GitHub Pages
- Try from a different network (mobile hotspot)
- Contact IT if using corporate network

### 4. Alternative Installation Methods

#### Method A: Developer Mode (Advanced)
1. Enable Developer Mode in Excel:
   - Go to **Excel** > **Preferences** > **Security & Privacy**
   - Check "Allow add-ins to start when Excel starts"
2. Use the manifest file directly

#### Method B: Office 365 Admin Center
1. If you have Office 365 admin access
2. Upload the manifest through the admin center
3. Deploy to your organization

### 5. Error Messages and Solutions

#### "Add-in failed to load"
- Check internet connection
- Verify GitHub Pages is accessible
- Try clearing Excel cache

#### "Invalid manifest"
- Download a fresh copy of manifest.xml
- Ensure file is not corrupted
- Check Excel version compatibility

#### "Add-in not found"
- Verify the manifest.xml file is in the correct location
- Check that all URLs in manifest are accessible
- Try the sideload package method

### 6. Testing the Add-in

Once installed:
1. Look for "Bill.com Data" button in Excel ribbon
2. Click the button to open the task pane
3. Enter your Bill.com credentials
4. Click "Fetch Bills" to test functionality

### 7. Still Having Issues?

If none of the above work:
1. Try on a different Mac/computer
2. Check Excel version and update if needed
3. Try from a different network
4. Contact support with specific error messages

## Files Available for Download

- **Manifest file**: https://nunoxsantos.github.io/excel-addin/manifest.xml
- **Sideload package**: https://github.com/nunoxsantos/excel-addin/raw/main/excel-addin-sideload.zip
- **Repository**: https://github.com/nunoxsantos/excel-addin
