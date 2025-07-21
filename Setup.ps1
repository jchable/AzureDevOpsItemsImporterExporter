# Script de configuration pour Azure DevOps Items Importer/Exporter
# Aide à la configuration initiale du script

Write-Host "=== Configuration Azure DevOps Items Importer/Exporter ===" -ForegroundColor Cyan
Write-Host ""

# Vérification de PowerShell
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -lt 5) {
    Write-Host "⚠️  PowerShell 5.1 ou supérieur requis. Version actuelle: $($psVersion.ToString())" -ForegroundColor Red
    exit 1
} else {
    Write-Host "✅ PowerShell $($psVersion.ToString()) détecté" -ForegroundColor Green
}

# Collecte des informations
Write-Host ""
Write-Host "📋 Informations requises :" -ForegroundColor Yellow
Write-Host ""

$Organization = Read-Host "Nom de l'organisation Azure DevOps"
$Project = Read-Host "Nom du projet"

Write-Host ""
Write-Host "🔑 Token d'accès personnel :" -ForegroundColor Yellow
Write-Host "1. Allez sur https://dev.azure.com/$Organization"
Write-Host "2. Cliquez sur votre profil → Personal Access Tokens"
Write-Host "3. Créez un token avec les permissions : Work Items (Read & Write), Project and Team (Read)"
Write-Host ""

$Token = Read-Host "Token d'accès (sera masqué)" -AsSecureString
$TokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token))

# Test de la connexion
Write-Host ""
Write-Host "🔍 Test de la connexion..." -ForegroundColor Yellow

try {
    $headers = @{
        Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$TokenPlain"))
        "Content-Type" = "application/json"
    }
    
    $uri = "https://dev.azure.com/$Organization/_apis/projects/$Project" + "?api-version=6.0"
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    
    Write-Host "✅ Connexion réussie au projet: $($response.name)" -ForegroundColor Green
    Write-Host "   Description: $($response.description)" -ForegroundColor Gray
    Write-Host "   URL: $($response.url)" -ForegroundColor Gray
}
catch {
    Write-Host "❌ Échec de la connexion: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Vérifiez :" -ForegroundColor Yellow
    Write-Host "- Le nom de l'organisation et du projet"
    Write-Host "- La validité du token d'accès"
    Write-Host "- Les permissions du token"
    exit 1
}

# Génération des exemples de commandes
Write-Host ""
Write-Host "📝 Commandes générées :" -ForegroundColor Yellow
Write-Host ""

# Commandes simplifiées sans token (utilisation du fichier .pat)
$exportCmdSimple = ".\AzureDevOpsItemsImporter.ps1 -Organization `"$Organization`" -Project `"$Project`" -Action Export -CsvFilePath `"export-$(Get-Date -Format 'yyyy-MM-dd').csv`""
$importCmdSimple = ".\AzureDevOpsItemsImporter.ps1 -Organization `"$Organization`" -Project `"$Project`" -Action Import -CsvFilePath `"import.csv`" -Verbose"

# Commandes avec token explicite
$exportCmd = ".\AzureDevOpsItemsImporter.ps1 -Organization `"$Organization`" -Project `"$Project`" -PersonalAccessToken `"$TokenPlain`" -Action Export -CsvFilePath `"export-$(Get-Date -Format 'yyyy-MM-dd').csv`""
$importCmd = ".\AzureDevOpsItemsImporter.ps1 -Organization `"$Organization`" -Project `"$Project`" -PersonalAccessToken `"$TokenPlain`" -Action Import -CsvFilePath `"import.csv`" -Verbose"

Write-Host "Export (avec fichier .pat) :" -ForegroundColor Cyan
Write-Host $exportCmdSimple -ForegroundColor White
Write-Host ""

Write-Host "Import (avec fichier .pat) :" -ForegroundColor Cyan
Write-Host $importCmdSimple -ForegroundColor White
Write-Host ""

Write-Host "Export (avec token explicite) :" -ForegroundColor Gray
Write-Host $exportCmd -ForegroundColor DarkGray
Write-Host ""

Write-Host "Import (avec token explicite) :" -ForegroundColor Gray
Write-Host $importCmd -ForegroundColor DarkGray
Write-Host ""

# Proposition de créer un fichier .pat
Write-Host "💾 Sauvegarde du token :" -ForegroundColor Yellow
$saveToken = Read-Host "Créer un fichier 'personal-access-token.pat' pour éviter de retaper le token ? (Y/n)"
if ($saveToken -ne 'n' -and $saveToken -ne 'N') {
    try {
        $TokenPlain | Out-File -FilePath "personal-access-token.pat" -Encoding UTF8 -NoNewline
        Write-Host "✅ Token sauvegardé dans 'personal-access-token.pat'" -ForegroundColor Green
        Write-Host "   Le fichier est automatiquement exclu de Git (.gitignore)" -ForegroundColor Gray
    }
    catch {
        Write-Host "❌ Erreur lors de la création du fichier .pat: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Sauvegarde de la configuration
$configFile = "config.ps1"
$configContent = @"
# Configuration Azure DevOps Items Importer/Exporter
# Généré le $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

`$Organization = "$Organization"
`$Project = "$Project"

# Note: Le token est maintenant stocké dans personal-access-token.pat
# Exemples d'utilisation (le script lira automatiquement le fichier .pat) :
# Export : .\AzureDevOpsItemsImporter.ps1 -Organization `$Organization -Project `$Project -Action Export
# Import : .\AzureDevOpsItemsImporter.ps1 -Organization `$Organization -Project `$Project -Action Import -CsvFilePath "import.csv"
"@

$saveConfig = Read-Host "Sauvegarder la configuration dans '$configFile' ? (y/N)"
if ($saveConfig -eq 'y' -or $saveConfig -eq 'Y') {
    $configContent | Out-File -FilePath $configFile -Encoding UTF8
    Write-Host "✅ Configuration sauvegardée dans $configFile" -ForegroundColor Green
}

Write-Host ""
Write-Host "🎉 Configuration terminée! Vous pouvez maintenant utiliser le script." -ForegroundColor Green
Write-Host ""
Write-Host "Prochaines étapes :" -ForegroundColor Yellow
Write-Host "1. Pour exporter : utilisez la commande Export ci-dessus (le token sera lu automatiquement depuis le fichier .pat)"
Write-Host "2. Pour importer : modifiez le CSV exporté et utilisez la commande Import"
Write-Host "3. Consultez le README.md pour plus de détails"
Write-Host ""
Write-Host "💡 Conseil : Avec le fichier .pat, vous n'avez plus besoin de spécifier le token à chaque fois !" -ForegroundColor Cyan
