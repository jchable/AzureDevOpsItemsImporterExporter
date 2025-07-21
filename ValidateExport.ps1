# Script de comparaison et validation des exports CSV
# Aide à valider que tous les work items sont bien exportés

param(
    [Parameter(Mandatory=$false)]
    [string]$CsvFile = "test-export.csv"
)

Write-Host "=== Validation de l'export CSV ===" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $CsvFile)) {
    Write-Host "❌ Fichier CSV non trouvé: $CsvFile" -ForegroundColor Red
    exit 1
}

try {
    # Lecture du CSV
    $csvData = Import-Csv -Path $CsvFile -Encoding UTF8
    
    Write-Host "📊 Statistiques de l'export:" -ForegroundColor Yellow
    Write-Host "   Total work items: $($csvData.Count)" -ForegroundColor Green
    Write-Host ""
    
    # Statistiques par type
    Write-Host "📋 Répartition par type:" -ForegroundColor Yellow
    $typeStats = $csvData | Group-Object Type | Sort-Object Count -Descending
    foreach ($type in $typeStats) {
        Write-Host "   $($type.Name): $($type.Count)" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Statistiques par statut
    Write-Host "🔄 Répartition par statut:" -ForegroundColor Yellow
    $stateStats = $csvData | Group-Object State | Sort-Object Count -Descending
    foreach ($state in $stateStats) {
        Write-Host "   $($state.Name): $($state.Count)" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Vérification des IDs
    Write-Host "🔍 Validation des IDs:" -ForegroundColor Yellow
    $ids = $csvData | ForEach-Object { [int]$_.Id } | Sort-Object
    $minId = $ids | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
    $maxId = $ids | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
    $uniqueIds = $ids | Select-Object -Unique
    
    Write-Host "   Plage d'IDs: $minId - $maxId" -ForegroundColor Gray
    Write-Host "   IDs uniques: $($uniqueIds.Count)" -ForegroundColor Gray
    
    if ($uniqueIds.Count -ne $csvData.Count) {
        Write-Host "   ⚠️  Attention: IDs dupliqués détectés!" -ForegroundColor Yellow
    } else {
        Write-Host "   ✅ Tous les IDs sont uniques" -ForegroundColor Green
    }
    Write-Host ""
    
    # Vérification des champs obligatoires
    Write-Host "✅ Validation des champs obligatoires:" -ForegroundColor Yellow
    $requiredFields = @('Id', 'Type', 'Title', 'State')
    $validationErrors = @()
    
    foreach ($field in $requiredFields) {
        $emptyCount = ($csvData | Where-Object { -not $_.$field -or $_.$field.Trim() -eq '' }).Count
        if ($emptyCount -gt 0) {
            $validationErrors += "   ❌ Champ '$field': $emptyCount valeurs vides"
        } else {
            Write-Host "   ✅ Champ '$field': Toutes les valeurs présentes" -ForegroundColor Green
        }
    }
    
    if ($validationErrors.Count -gt 0) {
        Write-Host "   Erreurs détectées:" -ForegroundColor Red
        foreach ($validationErr in $validationErrors) {
            Write-Host $validationErr -ForegroundColor Red
        }
    }
    Write-Host ""
    
    # Work items avec assignation
    $assignedItems = ($csvData | Where-Object { $_.AssignedTo -and $_.AssignedTo.Trim() -ne '' }).Count
    Write-Host "👤 Work items assignés: $assignedItems / $($csvData.Count) ($([Math]::Round($assignedItems / $csvData.Count * 100, 1))%)" -ForegroundColor Cyan
    
    # Work items avec charge restante
    $workItems = ($csvData | Where-Object { $_.RemainingWork -and $_.RemainingWork.Trim() -ne '' }).Count
    Write-Host "⏱️  Work items avec charge: $workItems / $($csvData.Count) ($([Math]::Round($workItems / $csvData.Count * 100, 1))%)" -ForegroundColor Cyan
    
    # Work items avec tags
    $taggedItems = ($csvData | Where-Object { $_.Tags -and $_.Tags.Trim() -ne '' }).Count
    Write-Host "🏷️  Work items avec tags: $taggedItems / $($csvData.Count) ($([Math]::Round($taggedItems / $csvData.Count * 100, 1))%)" -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "🎉 Validation terminée!" -ForegroundColor Green
    Write-Host "   Fichier: $CsvFile" -ForegroundColor Gray
    Write-Host "   Taille: $((Get-Item $CsvFile).Length / 1KB | ForEach-Object { '{0:N1}' -f $_ }) KB" -ForegroundColor Gray
    
}
catch {
    Write-Host "❌ Erreur lors de la validation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "💡 Pour comparer avec un export précédent:" -ForegroundColor Yellow
Write-Host "   Compare-Object (Import-Csv ancien.csv) (Import-Csv nouveau.csv) -Property Id" -ForegroundColor White
