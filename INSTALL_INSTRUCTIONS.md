# How to Install Your Word Add-in

## Method 1: Registry Method (Recommended if Upload option not visible)

### Step 1: Download the Manifest File
1. Go to: https://github.com/Sami9188/word-text-formatter/raw/main/manifest.xml
2. Save it to: `C:\WordAddin\manifest.xml`
3. Create the `C:\WordAddin` folder if it doesn't exist

### Step 2: Add to Registry (Choose your Office version)

**For Office 365 / Office 2016/2019/2021:**
1. Press `Win + R`
2. Type `regedit` and press Enter
3. Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`
4. Right-click `TrustedCatalogs` → New → Key
5. Name it: `TextFormatter`
6. Right-click `TextFormatter` → New → String Value
7. Name it: `CatalogUrl`
8. Double-click `CatalogUrl` and set value to: `file:///C:/WordAddin/`
9. Click OK

**For Office 2013:**
- Use `15.0` instead of `16.0` in the path above

### Step 3: Restart Word
- Close Word completely
- Reopen Word
- The add-in should appear in the **Home** tab ribbon

---

## Method 2: Using the Add-ins Dialog

1. In Word, go to **Insert** → **Get Add-ins**
2. Look for **"MY ADD-INS"** tab at the top
3. Click **"Upload My Add-in"** button
4. Browse and select your `manifest.xml` file
5. Click **Upload**

---

## Troubleshooting

**Add-in doesn't appear after registry method:**
- Make sure the path in registry uses forward slashes: `file:///C:/WordAddin/`
- Ensure `manifest.xml` is in that folder
- Check Office version (16.0 for Office 365, 15.0 for Office 2013)
- Restart Word completely

**Still not working?**
- Try moving manifest.xml to a simpler path without spaces
- Check that the manifest.xml URLs point to: `https://sami9188.github.io/word-text-formatter/`

