# Azure Icons Setup Guide

## Overview
The Building Block Generator now supports both Unicode symbol icons and actual Azure service icons. By default, it uses Unicode symbols (🌐, 🤖, 🗄️, etc.), but you can enhance it with official Azure icons for a more professional appearance.

## Using Official Azure Icons

### Option 1: Download from Microsoft
1. Visit the [Azure Architecture Icons](https://learn.microsoft.com/en-us/azure/architecture/icons/) page
2. Download the official Azure icon set
3. Create an `azure_icons` folder in your project directory
4. Extract the icons to this folder

### Option 2: Manual Icon Collection
Create an `azure_icons` folder and add icon files with these naming patterns:

#### Required Azure Service Icons:
- `azure-web-apps.svg` or `azure-web-apps.png`
- `azure-container-apps.svg` or `azure-container-apps.png`
- `azure-kubernetes-service.svg` or `aks.png`
- `azure-openai.svg` or `azure-openai.png`
- `microsoft-fabric.svg` or `fabric.png`
- `azure-databricks.svg` or `databricks.png`
- `azure-sql-database.svg` or `azure-sql.png`
- `azure-cosmos-db.svg` or `cosmosdb.png`
- `azure-data-factory.svg` or `data-factory.png`
- `azure-api-management.svg` or `apim.png`
- `azure-logic-apps.svg` or `logic-apps.png`
- `azure-service-bus.svg` or `service-bus.png`
- `azure-active-directory.svg` or `azure-ad.png`
- `azure-key-vault.svg` or `key-vault.png`
- `azure-application-gateway.svg` or `app-gateway.png`
- `azure-load-balancer.svg` or `load-balancer.png`
- `azure-virtual-network.svg` or `vnet.png`
- `azure-monitor.svg` or `monitor.png`

### Folder Structure
```
c:\Users\fnazaret\dev\bootcamp\
├── direct_generator.py
├── azure_icons/
│   ├── azure-web-apps.svg
│   ├── azure-openai.svg
│   ├── azure-sql-database.svg
│   ├── azure-cosmos-db.svg
│   ├── azure-databricks.svg
│   └── ... (other Azure service icons)
├── azure_icons_setup.md
└── ... (other files)
```

## How It Works

1. **Icon Loading Priority:**
   - First, the generator tries to load an actual icon file from the `azure_icons` folder
   - If found, it adds the image to the slide and uses clean service names
   - If not found, it falls back to Unicode symbols

2. **Naming Conventions:**
   The system looks for icons using multiple filename patterns:
   - `service-name-with-hyphens.svg/png`
   - `service_name_with_underscores.svg/png`
   - `servicename.svg/png`

3. **Icon Sizing:**
   - Actual Azure icons are displayed at 0.3" x 0.3"
   - Unicode symbols are embedded in text

## Testing Icon Setup

Run the generator to see which icons are being loaded:
```powershell
cd c:\Users\fnazaret\dev\bootcamp
python direct_generator.py
```

The console will show messages like:
- `"Loaded Azure icon: azure-web-apps.svg"` ✅ (actual icon loaded)
- `"No icon file found for Azure OpenAI. Using fallback."` ⚠️ (using Unicode symbol)

## Benefits of Using Actual Icons

1. **Professional Appearance:** Official Azure icons look more professional
2. **Brand Consistency:** Maintains Microsoft Azure visual standards
3. **Better Recognition:** Stakeholders immediately recognize Azure services
4. **Scalability:** Icons remain crisp at different presentation sizes

## Fallback Behavior

If no `azure_icons` folder exists or icons are missing:
- ✅ Generator continues to work normally
- ✅ Uses Unicode symbols as visual indicators
- ✅ No errors or crashes
- ⚠️ Slightly less professional appearance

This dual approach ensures your presentations always work, whether or not you have official Azure icons available.