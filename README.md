# Azure DevOps Items Importer/Exporter - Documentation

## Description
Advanced PowerShell script to export and import Azure DevOps work items (Epics, Features, PBIs, Tasks, User Stories, Bugs) to/from a CSV file with full support for creation and update (upsert), unlimited item retrieval, and advanced validation.

## ✨ New Features

### 🔥 Unlimited Export
- **Complete Retrieval**: No more 80-item limit! Retrieves **ALL** work items from the project.
- **Robust Method**: Advanced pagination system with fallback by ID ranges.
- **Optimized Performance**: Batch processing of 200 items with intelligent error handling.

### 🎯 Advanced Execution Modes
- **Simulation Mode** (`-WhatIf`): Test your imports without making any changes.
- **Force Mode** (`-Force`): Automatic import without confirmation.
- **Analysis-Only Mode** (`-ShowSummaryOnly`): Analyze the CSV without executing the import.

### 📊 Intelligent Preliminary Analysis
- **Detailed Summary**: Visual tables of planned operations (creations/updates).
- **Advanced Validation**: Complete CSV verification with error detection.
- **Flexible Decimal Format**: Supports both commas AND periods for numeric values.

## 🚀 Quick Start

**First time using it? Run the setup script:**

```powershell
.\Setup.ps1
```

This script will automatically guide you through configuring Azure DevOps and creating your first exports/imports.

## Prerequisites

### 1. Personal Access Token (PAT)
1. Log in to Azure DevOps.
2. Click on your profile → **Personal Access Tokens**.
3. Create a new token with the following permissions:
   - **Work Items**: Read & Write
   - **Project and Team**: Read
4. Copy the token (it will not be displayed again).

### 2. PowerShell
- PowerShell 5.1 or higher.
- `Invoke-RestMethod` module (included by default).

## Configuration

### Automatic Configuration (Recommended)
🚀 **Use the automatic configuration script for your first use:**

```powershell
.\Setup.ps1
```

This interactive script will guide you to:
- ✅ Verify PowerShell prerequisites.
- ✅ Test the connection to Azure DevOps.
- ✅ Automatically create the `.pat` file to store your token.
- ✅ Generate the `config.ps1` file with your parameters.
- ✅ Provide ready-to-use commands.

### Manual Configuration

### Required Parameters
- **Organization**: Name of your Azure DevOps organization.
- **Project**: Name of the project.
- **Action**: `Export` or `Import`.

### Optional Parameters
- **PersonalAccessToken**: Your access token (optional if `.pat` file is present).
- **CsvFilePath**: Path to the CSV file (default: `workitems.csv`).
- **Verbose**: Detailed operation logs.

### Access Token Management
The script automatically searches for the token in the following order:
1. **Parameter**: Token provided via `-PersonalAccessToken`.
2. **.pat File**: Searches in the current directory:
   - `personal-access-token.pat`
   - `.pat`

To create a `.pat` file:
```powershell
"your_token_here" | Out-File -FilePath "personal-access-token.pat" -Encoding UTF8
```

## Usage

### Export Work Items

#### Full Export (Recommended)
```powershell
# Export ALL work items from the project (no more 80-item limit)
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Export -CsvFilePath "full-export.csv"
```

#### Export with Token as Parameter
```powershell
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -PersonalAccessToken "your_token" -Action Export -CsvFilePath "export.csv"
```

### 🆕 Advanced Import with Execution Modes

#### 1. Simulation Mode (MANDATORY before real import)
```powershell
# ALWAYS test with -WhatIf before the real import
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Import -CsvFilePath "upsert.csv" -WhatIf
```

#### 2. Analysis-Only Mode
```powershell
# Analyze the CSV file without making any changes
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Import -CsvFilePath "upsert.csv" -ShowSummaryOnly
```

#### 3. Real Import with Confirmation
```powershell
# Import with manual confirmation
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Import -CsvFilePath "upsert.csv"
```

#### 4. Automatic Import (Force Mode)
```powershell
# Import without confirmation (use with caution!)
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Import -CsvFilePath "upsert.csv" -Force
```

### 🆕 Upsert Feature (Create + Update)

The **upsert** feature allows you to automatically create new work items or update existing work items in the same CSV file.

#### Automatic Visual Summary
The script now displays a detailed summary before each import:
```
=== IMPORT SUMMARY ===

📊 File Analysis: upsert.csv (25 work items)

🆕 PLANNED CREATIONS (3 work items):
┌─────────────────┬────────────────────────────────┬─────────────┬──────────────────────┐
│ Type            │ Title                          │ State       │ AssignedTo           │
├─────────────────┼────────────────────────────────┼─────────────┼──────────────────────┤
│ Task            │ New API Task                   │ New         │ dev@company.com      │
│ Bug             │ Fix login bug                  │ New         │ qa@company.com       │
│ Feature         │ New feature                    │ New         │ pm@company.com       │
└─────────────────┴────────────────────────────────┴─────────────┴──────────────────────┘

🔄 PLANNED UPDATES (2 work items):
┌──────┬──────────────┬────────────────────────────────┬─────────────┬──────────────────────┐
│ ID   │ Type         │ Title                          │ State       │ AssignedTo           │
├──────┼──────────────┼────────────────────────────────┼─────────────┼──────────────────────┤
│ 1001 │ Epic         │ Updated Epic                   │ Active      │ manager@company.com  │
│ 1002 │ Task         │ Updated Task                   │ Active      │ dev@company.com      │
└──────┴──────────────┴────────────────────────────────┴─────────────┴──────────────────────┘

✅ SUMMARY:
   • Creations: 3 work items
   • Updates: 2 work items
   • Errors: 0
   • Warnings: 0
```

#### CSV Format for Upsert
- **To CREATE**: Leave the `Id` field empty or set it to `0`.
- **To UPDATE**: Specify the `Id` of the existing work item.

```csv
Id,Type,Title,State,AssignedTo,Description
,Task,New Task,New,user@company.com,Automatically created task
0,Bug,New Bug,New,dev@company.com,Detected bug
1001,Epic,Updated Epic,Active,manager@company.com,Updated epic
```

📖 **See the [Complete Upsert Guide](GUIDE-UPSERT.md) for more details**

## CSV File Format

### Supported Columns
| Column       | Required | Description                              | Format                              |
|--------------|----------|------------------------------------------|-------------------------------------|
| Id           | ✅        | Numeric ID of the work item              | Integer or empty for creation       |
| Type         | ✅        | Epic, Feature, Product Backlog Item, etc.| Text                                |
| Title        | ✅        | Title of the work item                   | Text                                |
| State        | ✅        | Status (New, Active, Resolved, etc.)     | Text                                |
| AssignedTo   | ❌        | Assigned person (full name or email)     | Text                                |
| Description  | ❌        | Detailed description                     | Text (HTML automatically converted) |
| RemainingWork| ❌        | Remaining work                           | **Decimal (comma OR period accepted)** |
| Tags         | ❌        | Tags separated by semicolons             | Text                                |
| IterationPath| ❌        | Iteration path                           | Text                                |
| AreaPath     | ❌        | Area path                                | Text                                |

### 🔢 Decimal Format Support
The script now supports **both formats** for numeric values:
- **French Format**: `0,5` `2,75` `1,25`
- **English Format**: `0.5` `2.75` `1.25`

Conversion is automatic to the format required by the Azure DevOps API.

### CSV Example
```csv
Id,Type,Title,State,AssignedTo,Description,RemainingWork,Tags,IterationPath
1234,Task,Implement API,Active,john.doe@company.com,Develop the REST API,0.5,backend;api,Project\Sprint 1
5678,Bug,Fix login bug,New,jane.smith@company.com,Login not working,2.25,bug;urgent,Project\Sprint 1
,Feature,New Feature,New,pm@company.com,Feature to be created,8,new;feature,Project\Sprint 2
0,Epic,New Epic,New,manager@company.com,Epic to be created,,,Project\Release 1
```

**Note**: Values like `0,5` and `2,25` are automatically converted to `0.5` and `2.25` for the Azure DevOps API.

## Validation and Security

### Advanced Automatic Validation
- ✅ Verifies connection to Azure DevOps.
- ✅ Complete CSV format validation with line-by-line analysis.
- ✅ Checks required and optional columns.
- ✅ Validates data types (numbers, email formats, etc.).
- ✅ **Supports both French (comma) and English (period) decimal formats.**
- ✅ Checks the existence of work items before modification.
- ✅ Full stop on error with detailed diagnostics.

### Validation Modes
1. **Preliminary Analysis**: CSV validation + visual summary.
2. **Simulation Mode** (`-WhatIf`): Full test without modification.
3. **Analysis-Only Mode** (`-ShowSummaryOnly`): Analysis without execution.

### Robust Error Handling
- If **even one error** is detected in the CSV, the import is **completely canceled**.
- Detailed display of all errors found with line numbers.
- Color-coded logs for easier tracking and debugging.
- Precise error messages with correction suggestions.

## Examples of Use

### 0. Initial Configuration (First Use)
```powershell
# First, run the setup script
.\Setup.ps1
```

### 1. Full Export (Recommended)
```powershell
# Export ALL work items (no more 80-item limit)
.\AzureDevOpsItemsImporter.ps1 -Organization "contoso" -Project "WebApp" -Action Export -CsvFilePath "full-export.csv"
```

### 2. Import with Preliminary Analysis
```powershell
# Analysis only (no modification)
.\AzureDevOpsItemsImporter.ps1 -Organization "contoso" -Project "WebApp" -Action Import -CsvFilePath "modifications.csv" -ShowSummaryOnly

# Test in simulation
.\AzureDevOpsItemsImporter.ps1 -Organization "contoso" -Project "WebApp" -Action Import -CsvFilePath "modifications.csv" -WhatIf

# Real import with confirmation
.\AzureDevOpsItemsImporter.ps1 -Organization "contoso" -Project "WebApp" -Action Import -CsvFilePath "modifications.csv" -Verbose
```

### 3. Automated Import
```powershell
# Import without confirmation (for automated scripts)
.\AzureDevOpsItemsImporter.ps1 -Organization "contoso" -Project "WebApp" -Action Import -CsvFilePath "modifications.csv" -Force
```

### 3. Creating a .pat File
```powershell
# Create the file with your token
"ghp_1234567890abcdef..." | Out-File -FilePath "personal-access-token.pat" -Encoding UTF8 -NoNewline

# Verify the file was created
Get-Content "personal-access-token.pat"
```

## Security

### Token Protection
- ✅ `.pat` files are automatically excluded by `.gitignore`.
- ✅ Never commit tokens to source code.
- ✅ Use tokens with minimal required permissions.
- ✅ Regularly regenerate tokens.

### Best Practices
- Store your tokens in local `.pat` files.
- Use explicit file names (e.g., `azure-devops-prod.pat`).
- Set an expiration date on your tokens.
- Revoke unused tokens.

## Limitations and Constraints

### Technical Limitations
- HTML descriptions are converted to plain text during export.
- Only standard work item types are supported (Epic, Feature, PBI, Task, User Story, Bug).
- Attachments are not managed.

### Performance Constraints
- ✅ **No limit on the number of exported work items** (80-item limit fixed).
- Batch processing of 200 work items for export.
- Automatic pause every 10 batches to avoid API overload.
- Fallback method by ID ranges in case of API issues.

### Security Constraints
- PAT tokens must have Work Items (Read & Write) and Project (Read) permissions.
- Secure storage of tokens in `.pat` files excluded from Git.

## Utility Scripts

### 🔧 Setup.ps1
Interactive configuration script for first use:
```powershell
.\Setup.ps1
```

### ✅ ValidateExport.ps1
Validation script to analyze an exported CSV file:
```powershell
.\ValidateExport.ps1 -CsvFile "export.csv"
```

### 📖 docs/GUIDE-UPSERT.md
Complete guide to using the upsert feature (create + update):
- CSV format for creation and update.
- Practical usage examples.
- Best practices and troubleshooting.

## Troubleshooting

### Import Issues

#### Recommended Test Mode
```powershell
# ALWAYS use -WhatIf before the real import
.\AzureDevOpsItemsImporter.ps1 -Organization "myorg" -Project "myproject" -Action Import -CsvFilePath "file.csv" -WhatIf
```

### Token Error
- Verify that the `.pat` file exists and contains the correct token.
- Confirm that the token has not expired.
- Ensure the token has the correct permissions.

### Connection Error
- Verify your access token.
- Confirm the organization and project names (case-sensitive).
- Ensure the token has the correct permissions.
- Use the test script: `.\TestConnection.ps1 -Organization "myorg" -Project "myproject" -Token "mytoken"`.
