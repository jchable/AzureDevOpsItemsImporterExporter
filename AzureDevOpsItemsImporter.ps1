# Azure DevOps Items Importer/Exporter
# Script pour exporter et importer des work items Azure DevOps

param(
    [Parameter(Mandatory=$true)]
    [string]$Organization,
    
    [Parameter(Mandatory=$true)]
    [string]$Project,
    
    [Parameter(Mandatory=$false)]
    [string]$PersonalAccessToken,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("Export", "Import")]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$CsvFilePath = "workitems.csv",
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowSummaryOnly
)

# Configuration globale
$ErrorActionPreference = "Stop"

# Fonction pour écrire des logs avec couleurs
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        "Info"    { Write-Host "[$timestamp] [INFO] $Message" -ForegroundColor Cyan }
        "Warning" { Write-Host "[$timestamp] [WARNING] $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[$timestamp] [ERROR] $Message" -ForegroundColor Red }
        "Success" { Write-Host "[$timestamp] [SUCCESS] $Message" -ForegroundColor Green }
    }
}

# Fonction pour obtenir le token d'accès
function Get-PersonalAccessToken {
    param([string]$ProvidedToken)
    
    # Si un token est fourni en paramètre, l'utiliser
    if ($ProvidedToken) {
        Write-Log "Utilisation du token fourni en paramètre" "Info"
        return $ProvidedToken
    }
    
    # Chercher un fichier .pat dans le répertoire courant
    $patFiles = @(
        ".\personal-access-token.pat",
        ".\token.pat",
        ".\pat.token",
        ".\.pat"
    )
    
    foreach ($patFile in $patFiles) {
        if (Test-Path $patFile) {
            try {
                Write-Log "Token trouvé dans le fichier: $patFile" "Info"
                $token = Get-Content $patFile -Raw -Encoding UTF8
                return $token.Trim()
            }
            catch {
                Write-Log "Erreur lors de la lecture du fichier $patFile : $($_.Exception.Message)" "Warning"
                continue
            }
        }
    }
    
    # Si aucun fichier trouvé, demander le token
    Write-Log "Aucun fichier .pat trouvé. Fichiers recherchés:" "Warning"
    foreach ($patFile in $patFiles) {
        Write-Log "  - $patFile" "Warning"
    }
    
    throw "Token d'accès personnel requis. Fournissez-le via le paramètre -PersonalAccessToken ou créez un fichier .pat (ex: personal-access-token.pat) dans le répertoire courant."
}

# Obtention du token
$Token = Get-PersonalAccessToken -ProvidedToken $PersonalAccessToken

$Global:Headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
    "Content-Type" = "application/json"
}
$Global:BaseUri = "https://dev.azure.com/$Organization/_apis"

# Fonction pour tester la connexion Azure DevOps
function Test-AzureDevOpsConnection {
    try {
        Write-Log "Test de la connexion à Azure DevOps..." "Info"
        $uri = "$Global:BaseUri/projects/$Project" + "?api-version=6.0"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Global:Headers
        Write-Log "Connexion réussie au projet: $($response.name)" "Success"
        return $true
    }
    catch {
        Write-Log "Échec de la connexion: $($_.Exception.Message)" "Error"
        return $false
    }
}

# Fonction pour récupérer tous les work items
function Get-AllWorkItems {
    try {
        Write-Log "Récupération de tous les work items..." "Info"
        
        # Première requête pour obtenir le nombre total de work items
        $countQuery = @{
            query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$Project' AND [System.WorkItemType] IN ('Epic', 'Feature', 'Product Backlog Item', 'Task', 'User Story', 'Bug')"
        } | ConvertTo-Json

        $wiqlUri = "$Global:BaseUri/wit/wiql?api-version=6.0"
        $countResult = Invoke-RestMethod -Uri $wiqlUri -Method Post -Body $countQuery -Headers $Global:Headers
        
        Write-Log "Nombre total de work items dans le projet: $($countResult.workItems.Count)" "Info"
        
        if ($countResult.workItems.Count -eq 0) {
            Write-Log "Aucun work item trouvé dans le projet." "Warning"
            return @()
        }

        # CORRECTION: Toujours utiliser la méthode par batch pour garantir la récupération de TOUS les éléments
        # La requête WIQL standard a une limite cachée à ~200 résultats même avec $top
        Write-Log "Utilisation de la méthode par batch pour garantir la récupération complète..." "Info"
        return Get-AllWorkItemsLargeBatch

    }
    catch {
        Write-Log "Erreur lors de la récupération des work items: $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour récupérer tous les work items en cas de gros volume (méthode alternative)
function Get-AllWorkItemsLargeBatch {
    try {
        Write-Log "Récupération par méthode de pagination pour récupération complète..." "Info"
        
        $allWorkItems = @()
        $allWorkItemIds = @()
        
        # Utiliser une approche différente : récupérer TOUS les IDs d'abord sans limite
        # puis les traiter par batch
        Write-Log "Étape 1: Récupération de tous les IDs de work items..." "Info"
        
        # Requête WIQL pour obtenir TOUS les IDs sans pagination
        $wiqlQuery = @{
            query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$Project' AND [System.WorkItemType] IN ('Epic', 'Feature', 'Product Backlog Item', 'Task', 'User Story', 'Bug') ORDER BY [System.Id]"
        } | ConvertTo-Json

        # Utiliser l'API WIQL avec un $top très élevé pour contourner la limite
        $wiqlUri = "$Global:BaseUri/wit/wiql?`$top=20000&api-version=6.0"
        
        try {
            $queryResult = Invoke-RestMethod -Uri $wiqlUri -Method Post -Body $wiqlQuery -Headers $Global:Headers
            $allWorkItemIds = $queryResult.workItems | ForEach-Object { $_.id }
            Write-Log "IDs récupérés: $($allWorkItemIds.Count) work items trouvés" "Success"
        }
        catch {
            Write-Log "La requête avec top=20000 a échoué, utilisation de la méthode alternative..." "Warning"
            
            # Méthode alternative : récupération par plages d'IDs
            $allWorkItemIds = Get-WorkItemIdsByRange -Project $Project
        }
        
        if ($allWorkItemIds.Count -eq 0) {
            Write-Log "Aucun work item trouvé." "Warning"
            return @()
        }
        
        # Étape 2: Récupérer les détails par batch
        Write-Log "Étape 2: Récupération des détails pour $($allWorkItemIds.Count) work items..." "Info"
        
        $batchSize = 200
        $totalBatches = [Math]::Ceiling($allWorkItemIds.Count / $batchSize)
        $currentBatch = 0
        
        for ($i = 0; $i -lt $allWorkItemIds.Count; $i += $batchSize) {
            $currentBatch++
            $batch = $allWorkItemIds[$i..([Math]::Min($i + $batchSize - 1, $allWorkItemIds.Count - 1))]
            $idsString = $batch -join ","
            
            Write-Log "  Batch $currentBatch/$totalBatches : IDs $($batch[0]) à $($batch[-1]) ($($batch.Count) items)" "Info"
            
            $detailsUri = "$Global:BaseUri/wit/workitems?ids=$idsString&`$expand=all&api-version=6.0"
            
            try {
                $batchResult = Invoke-RestMethod -Uri $detailsUri -Method Get -Headers $Global:Headers
                $allWorkItems += $batchResult.value
                
                $percentComplete = [Math]::Round(($currentBatch / $totalBatches) * 100, 1)
                Write-Progress -Activity "Récupération des work items" -Status "Batch $currentBatch/$totalBatches - Total récupéré: $($allWorkItems.Count) work items" -PercentComplete $percentComplete
            }
            catch {
                Write-Log "Erreur lors de la récupération du batch $currentBatch : $($_.Exception.Message)" "Error"
                # Continuer avec le batch suivant en cas d'erreur
                continue
            }
            
            # Petite pause pour éviter de surcharger l'API
            if ($currentBatch % 10 -eq 0) {
                Start-Sleep -Milliseconds 500
            }
        }
        
        Write-Progress -Activity "Récupération des work items" -Completed
        Write-Log "🎉 Récupération complète terminée: $($allWorkItems.Count) work items traités au total." "Success"
        
        return $allWorkItems
    }
    catch {
        Write-Log "Erreur lors de la récupération étendue des work items: $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour convertir les work items en format CSV
function Convert-WorkItemsToCsv {
    param([array]$WorkItems)
    
    Write-Log "Conversion des work items en format CSV..." "Info"
    
    $csvData = @()
    foreach ($item in $WorkItems) {
        $fields = $item.fields
        
        $csvRow = [PSCustomObject]@{
            Id = $item.id
            Type = $fields.'System.WorkItemType'
            Title = $fields.'System.Title'
            State = $fields.'System.State'
            AssignedTo = if ($fields.'System.AssignedTo') { $fields.'System.AssignedTo'.displayName } else { "" }
            Description = if ($fields.'System.Description') { $fields.'System.Description' -replace '<[^>]+>', '' -replace '\r?\n', ' ' } else { "" }
            RemainingWork = if ($fields.'Microsoft.VSTS.Scheduling.RemainingWork') { $fields.'Microsoft.VSTS.Scheduling.RemainingWork' } else { "" }
            Tags = if ($fields.'System.Tags') { $fields.'System.Tags' } else { "" }
            IterationPath = if ($fields.'System.IterationPath') { $fields.'System.IterationPath' } else { "" }
            AreaPath = if ($fields.'System.AreaPath') { $fields.'System.AreaPath' } else { "" }
        }
        $csvData += $csvRow
    }
    
    Write-Log "Conversion terminée: $($csvData.Count) lignes créées." "Success"
    return $csvData
}

# Fonction pour exporter vers CSV
function Export-WorkItemsToCsv {
    param(
        [string]$FilePath
    )
    
    try {
        $workItems = Get-AllWorkItems
        if ($workItems.Count -eq 0) {
            Write-Log "Aucun work item à exporter." "Warning"
            return
        }
        
        $csvData = Convert-WorkItemsToCsv -WorkItems $workItems
        $csvData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        
        Write-Log "Export réussi vers: $FilePath" "Success"
        Write-Log "Nombre de work items exportés: $($csvData.Count)" "Info"
    }
    catch {
        Write-Log "Erreur lors de l'export: $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour valider le fichier CSV avant import
function Test-CsvFile {
    param([string]$FilePath)
    
    Write-Log "Validation du fichier CSV: $FilePath" "Info"
    
    if (-not (Test-Path $FilePath)) {
        throw "Le fichier CSV n'existe pas: $FilePath"
    }
    
    try {
        $csvData = Import-Csv -Path $FilePath -Encoding UTF8
    }
    catch {
        throw "Impossible de lire le fichier CSV: $($_.Exception.Message)"
    }
    
    if ($csvData.Count -eq 0) {
        throw "Le fichier CSV est vide."
    }
    
    # Validation des colonnes requises
    # Note: Id peut être vide pour création, mais Type et Title sont obligatoires
    $requiredColumns = @('Type', 'Title')
    $recommendedColumns = @('Id', 'State', 'AreaPath')
    $missingColumns = @()
    $missingRecommended = @()
    
    foreach ($column in $requiredColumns) {
        if (-not ($csvData[0].PSObject.Properties.Name -contains $column)) {
            $missingColumns += $column
        }
    }
    
    foreach ($column in $recommendedColumns) {
        if (-not ($csvData[0].PSObject.Properties.Name -contains $column)) {
            $missingRecommended += $column
        }
    }
    
    if ($missingColumns.Count -gt 0) {
        throw "Colonnes obligatoires manquantes dans le CSV: $($missingColumns -join ', ')"
    }
    
    if ($missingRecommended.Count -gt 0) {
        Write-Log "Colonnes recommandées manquantes: $($missingRecommended -join ', ')" "Warning"
    }
    
    # Validation des données ligne par ligne
    $errors = @()
    for ($i = 0; $i -lt $csvData.Count; $i++) {
        $row = $csvData[$i]
        $lineNumber = $i + 2  # +2 car ligne 1 = headers et index commence à 0
        
        # Validation ID (peut être vide pour création, 0 pour création, ou un nombre positif pour mise à jour)
        if ($row.Id -and $row.Id.Trim() -ne '' -and $row.Id -notmatch '^\d+$') {
            $errors += "Ligne $lineNumber : ID invalide: '$($row.Id)'. Doit être un nombre positif ou vide pour création."
        }
        
        # Validation Type
        $validTypes = @('Epic', 'Feature', 'Product Backlog Item', 'Task', 'User Story', 'Bug')
        if (-not $row.Type -or $row.Type -notin $validTypes) {
            $errors += "Ligne $lineNumber : Type invalide: '$($row.Type)'. Types valides: $($validTypes -join ', ')"
        }
        
        # Validation Title
        if (-not $row.Title -or $row.Title.Trim() -eq '') {
            $errors += "Ligne $lineNumber : Titre manquant ou vide"
        }
        
        # Validation RemainingWork (si présent)
        if ($row.RemainingWork -and $row.RemainingWork -ne '' -and $row.RemainingWork -notmatch '^\d*[,.]?\d*$') {
            $errors += "Ligne $lineNumber : Charge restante invalide: '$($row.RemainingWork)'. Doit être un nombre (avec point ou virgule comme séparateur décimal)."
        }
    }
    
    if ($errors.Count -gt 0) {
        Write-Log "Erreurs de validation détectées:" "Error"
        foreach ($err in $errors) {
            Write-Log $err "Error"
        }
        throw "Le fichier CSV contient $($errors.Count) erreur(s). Import annulé."
    }
    
    Write-Log "Validation du CSV réussie: $($csvData.Count) lignes valides." "Success"
    return $csvData
}

# Fonction pour mettre à jour un work item
function Update-WorkItem {
    param(
        [int]$Id,
        [hashtable]$Fields
    )
    
    $updateOperations = @()
    
    foreach ($fieldName in $Fields.Keys) {
        if ($null -ne $Fields[$fieldName] -and $Fields[$fieldName] -ne '') {
            $updateOperations += @{
                op = "replace"
                path = "/fields/$fieldName"
                value = $Fields[$fieldName]
            }
        }
    }
    
    if ($updateOperations.Count -eq 0) {
        Write-Log "Aucune mise à jour nécessaire pour le work item $Id" "Info"
        return
    }
    
    $updateBody = $updateOperations | ConvertTo-Json -Depth 3
    $updateUri = "$Global:BaseUri/wit/workitems/$Id" + "?api-version=6.0"
    
    $headers = $Global:Headers.Clone()
    $headers["Content-Type"] = "application/json-patch+json"
    
    try {
        if ($WhatIf) {
            Write-Log "[SIMULATION] Mise à jour du work item $Id avec $($updateOperations.Count) champs" "Info"
            return @{ id = $Id; fields = $Fields }
        }
        
        $result = Invoke-RestMethod -Uri $updateUri -Method Patch -Body $updateBody -Headers $headers
        Write-Log "Work item $Id mis à jour avec succès" "Success"
        return $result
    }
    catch {
        Write-Log "Erreur lors de la mise à jour du work item $Id : $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour créer un nouveau work item
function New-WorkItem {
    param(
        [string]$Type,
        [hashtable]$Fields
    )
    
    $createOperations = @()
    
    # Ajouter tous les champs fournis
    foreach ($fieldName in $Fields.Keys) {
        if ($null -ne $Fields[$fieldName] -and $Fields[$fieldName] -ne '') {
            $createOperations += @{
                op = "add"
                path = "/fields/$fieldName"
                value = $Fields[$fieldName]
            }
        }
    }
    
    if ($createOperations.Count -eq 0) {
        throw "Aucun champ fourni pour créer le work item"
    }
    
    $createBody = $createOperations | ConvertTo-Json -Depth 3
    $createUri = "$Global:BaseUri/wit/workitems/`$$Type" + "?api-version=6.0"
    
    $headers = $Global:Headers.Clone()
    $headers["Content-Type"] = "application/json-patch+json"
    
    try {
        if ($WhatIf) {
            Write-Log "[SIMULATION] Création d'un work item de type $Type avec $($createOperations.Count) champs" "Info"
            return @{ id = -1; fields = $Fields }
        }
        
        $result = Invoke-RestMethod -Uri $createUri -Method Post -Body $createBody -Headers $headers
        Write-Log "Work item créé avec succès: ID $($result.id), Type: $Type" "Success"
        return $result
    }
    catch {
        Write-Log "Erreur lors de la création du work item de type $Type : $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour vérifier si un work item existe
function Test-WorkItemExists {
    param([int]$Id)
    
    if ($WhatIf) {
        # En mode WhatIf, on simule que les IDs > 0 existent
        return ($Id -gt 0)
    }
    
    try {
        $uri = "$Global:BaseUri/wit/workitems/$Id" + "?api-version=6.0"
        $result = Invoke-RestMethod -Uri $uri -Method Get -Headers $Global:Headers
        return $true
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            return $false
        }
        # Pour les autres erreurs, on relance l'exception
        throw
    }
}

# Fonction pour faire un upsert (update ou insert) d'un work item
function Upsert-WorkItem {
    param(
        [string]$Id,  # Peut être vide pour création
        [string]$Type,
        [hashtable]$Fields
    )
    
    # Si l'ID est fourni et n'est pas vide, essayer de mettre à jour
    if ($Id -and $Id.Trim() -ne '' -and $Id -match '^\d+$') {
        $workItemId = [int]$Id
        
        if (Test-WorkItemExists -Id $workItemId) {
            Write-Log "Work item $workItemId existe - Mise à jour" "Info"
            return Update-WorkItem -Id $workItemId -Fields $Fields
        } else {
            Write-Log "Work item $workItemId n'existe pas - Création d'un nouveau work item" "Warning"
        }
    }
    
    # Créer un nouveau work item
    if (-not $Type -or $Type.Trim() -eq '') {
        throw "Le type de work item est requis pour la création"
    }
    
    Write-Log "Création d'un nouveau work item de type: $Type" "Info"
    return New-WorkItem -Type $Type -Fields $Fields
}

# Fonction pour importer depuis CSV
function Import-WorkItemsFromCsv {
    param([string]$FilePath)
    
    try {
        # 1. Analyse préliminaire des opérations
        Write-Log "Analyse préliminaire du fichier CSV..." "Info"
        $analysis = Analyze-ImportOperations -FilePath $FilePath
        
        # 2. Affichage du récapitulatif
        Show-ImportSummary -Analysis $analysis
        
        # 3. Mode ShowSummaryOnly - arrêt après affichage
        if ($ShowSummaryOnly) {
            Write-Log "Mode affichage seulement - Analyse terminée" "Info"
            return
        }
        
        # 4. Demande de confirmation (sauf en mode WhatIf ou Force)
        if (-not (Confirm-ImportExecution -Analysis $analysis -Force:$Force -WhatIf:$WhatIf)) {
            Write-Log "Import annulé" "Info"
            return
        }
        
        # 5. Exécution de l'import
        Write-Log "Début de l'exécution de l'import..." "Info"
        
        # Recharger les données CSV pour l'exécution
        $csvData = Test-CsvFile -FilePath $FilePath
        
        $successCount = 0
        $errorCount = 0
        $createdCount = 0
        $updatedCount = 0
        
        foreach ($row in $csvData) {
            try {
                Write-Progress -Activity "Import des work items" -Status "Traitement du work item $($row.Id)" -PercentComplete (($successCount + $errorCount) / $csvData.Count * 100)
                
                # Préparation des champs à mettre à jour
                $fieldsToUpdate = @{}
                
                if ($row.Title) { $fieldsToUpdate['System.Title'] = $row.Title }
                if ($row.State) { $fieldsToUpdate['System.State'] = $row.State }
                if ($row.Description) { $fieldsToUpdate['System.Description'] = $row.Description }
                if ($row.RemainingWork -and $row.RemainingWork -ne '') { 
                    # Convertir les virgules en points pour le format API Azure DevOps
                    $remainingWorkValue = $row.RemainingWork.Replace(',', '.')
                    $fieldsToUpdate['Microsoft.VSTS.Scheduling.RemainingWork'] = [double]$remainingWorkValue 
                }
                if ($row.Tags) { $fieldsToUpdate['System.Tags'] = $row.Tags }
                if ($row.IterationPath) { $fieldsToUpdate['System.IterationPath'] = $row.IterationPath }
                if ($row.AreaPath) { $fieldsToUpdate['System.AreaPath'] = $row.AreaPath }
                if ($row.AssignedTo) { $fieldsToUpdate['System.AssignedTo'] = $row.AssignedTo }
                
                # Déterminer s'il s'agit d'une création ou d'une mise à jour
                $isCreation = $false
                if (-not $row.Id -or $row.Id.Trim() -eq '' -or $row.Id -eq '0') {
                    $isCreation = $true
                } elseif ($row.Id -match '^\d+$') {
                    # Vérifier si le work item existe
                    if (-not (Test-WorkItemExists -Id ([int]$row.Id))) {
                        $isCreation = $true
                    }
                }
                
                # Upsert du work item
                $result = Upsert-WorkItem -Id $row.Id -Type $row.Type -Fields $fieldsToUpdate
                
                if ($isCreation) {
                    $createdCount++
                } else {
                    $updatedCount++
                }
                $successCount++
                
                # Utilisation du paramètre Verbose standard de PowerShell
                Write-Verbose "Work item $($row.Id) traité avec succès"
            }
            catch {
                $errorCount++
                Write-Log "Erreur pour le work item $($row.Id): $($_.Exception.Message)" "Error"
            }
        }
        
        Write-Progress -Activity "Import des work items" -Completed
        
        Write-Log "Import terminé. Succès: $successCount (Créés: $createdCount, Mis à jour: $updatedCount), Erreurs: $errorCount" "Success"
        
        if ($errorCount -gt 0) {
            Write-Log "Des erreurs sont survenues durant l'import. Vérifiez les logs ci-dessus." "Warning"
        }
    }
    catch {
        Write-Log "Erreur critique durant l'import: $($_.Exception.Message)" "Error"
        throw
    }
}

# Fonction pour analyser les opérations d'import sans les exécuter
function Analyze-ImportOperations {
    param([string]$FilePath)
    
    try {
        # Validation du fichier CSV
        $csvData = Test-CsvFile -FilePath $FilePath
        
        $analysis = @{
            TotalItems = $csvData.Count
            CreationItems = @()
            UpdateItems = @()
            Errors = @()
            Warnings = @()
            FilePath = $FilePath
        }
        
        Write-Log "Analyse préliminaire de $($csvData.Count) work items..." "Info"
        
        foreach ($row in $csvData) {
            try {
                $operation = @{
                    RowNumber = $csvData.IndexOf($row) + 2  # +2 car ligne 1 = headers et index commence à 0
                    Id = $row.Id
                    Type = $row.Type
                    Title = $row.Title
                    State = $row.State
                    AssignedTo = $row.AssignedTo
                    Fields = @{
                        # Liste des champs à inclure dans l'analyse
                        'System.Title' = $row.Title
                        'System.State' = $row.State
                        'System.Description' = $row.Description
                        'Microsoft.VSTS.Scheduling.RemainingWork' = if ($row.RemainingWork -and $row.RemainingWork -ne '') { [double]($row.RemainingWork.Replace(',', '.')) } else { $null }
                        'System.Tags' = $row.Tags
                        'System.IterationPath' = $row.IterationPath
                        'System.AreaPath' = $row.AreaPath
                        'System.AssignedTo' = $row.AssignedTo
                    }
                }
                
                # Déterminer le type d'opération
                $isCreation = $false
                if (-not $row.Id -or $row.Id.Trim() -eq '' -or $row.Id -eq '0') {
                    $isCreation = $true
                    $operation.OperationType = "Create"
                } elseif ($row.Id -match '^\d+$') {
                    $workItemId = [int]$row.Id
                    if (Test-WorkItemExists -Id $workItemId) {
                        $operation.OperationType = "Update"
                        # Ajouter un avertissement pour les modifications
                        $analysis.Warnings += "Work item $workItemId sera modifié"
                    } else {
                        $isCreation = $true
                        $operation.OperationType = "Create"
                        $analysis.Warnings += "Work item $workItemId n'existe pas, sera créé comme nouveau work item"
                    }
                }
                
                if ($isCreation) {
                    $analysis.CreationItems += $operation
                } else {
                    $analysis.UpdateItems += $operation
                }
                
            }
            catch {
                $analysis.Errors += "Ligne $($csvData.IndexOf($row) + 2): $($_.Exception.Message)"
            }
        }
        
        Write-Log "Analyse terminée: $($analysis.CreationItems.Count) créations, $($analysis.UpdateItems.Count) modifications" "Info"
        return $analysis
    }
    catch {
        throw "Erreur lors de l'analyse: $($_.Exception.Message)"
    }
}

# Fonction pour afficher le récapitulatif des opérations d'import
function Show-ImportSummary {
    param([hashtable]$Analysis)
    
    Write-Host ""
    Write-Host "=== RÉCAPITULATIF DE L'IMPORT ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "📊 Analyse du fichier: $([System.IO.Path]::GetFileName($Analysis.FilePath)) ($($Analysis.TotalItems) work items)" -ForegroundColor Yellow
    Write-Host ""
    
    # Affichage des créations
    if ($Analysis.CreationItems.Count -gt 0) {
        Write-Host "🆕 CRÉATIONS PRÉVUES ($($Analysis.CreationItems.Count) work items):" -ForegroundColor Green
        Write-Host "┌─────────────────┬────────────────────────────────┬─────────────┬──────────────────────┐" -ForegroundColor Gray
        Write-Host "│ Type            │ Title                          │ State       │ AssignedTo           │" -ForegroundColor Gray
        Write-Host "├─────────────────┼────────────────────────────────┼─────────────┼──────────────────────┤" -ForegroundColor Gray
        
        foreach ($item in $Analysis.CreationItems) {
            $type = $item.Type.PadRight(15)
            $title = if ($item.Title.Length -gt 30) { $item.Title.Substring(0, 27) + "..." } else { $item.Title.PadRight(30) }
            $state = $item.State.PadRight(11)
            $assignedTo = if ($item.AssignedTo) { 
                if ($item.AssignedTo.Length -gt 20) { $item.AssignedTo.Substring(0, 17) + "..." } else { $item.AssignedTo.PadRight(20) }
            } else { "".PadRight(20) }
            
            Write-Host "│ $type │ $title │ $state │ $assignedTo │" -ForegroundColor White
        }
        Write-Host "└─────────────────┴────────────────────────────────┴─────────────┴──────────────────────┘" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Affichage des modifications
    if ($Analysis.UpdateItems.Count -gt 0) {
        Write-Host "🔄 MODIFICATIONS PRÉVUES ($($Analysis.UpdateItems.Count) work items):" -ForegroundColor Yellow
        Write-Host "┌──────┬──────────────┬────────────────────────────────┬─────────────┬──────────────────────┐" -ForegroundColor Gray
        Write-Host "│ ID   │ Type         │ Title                          │ State       │ AssignedTo           │" -ForegroundColor Gray
        Write-Host "├──────┼──────────────┼────────────────────────────────┼─────────────┼──────────────────────┤" -ForegroundColor Gray
        
        foreach ($item in $Analysis.UpdateItems) {
            $id = $item.Id.ToString().PadRight(4)
            $type = $item.Type.PadRight(12)
            $title = if ($item.Title.Length -gt 30) { $item.Title.Substring(0, 27) + "..." } else { $item.Title.PadRight(30) }
            $state = $item.State.PadRight(11)
            $assignedTo = if ($item.AssignedTo) { 
                if ($item.AssignedTo.Length -gt 20) { $item.AssignedTo.Substring(0, 17) + "..." } else { $item.AssignedTo.PadRight(20) }
            } else { "".PadRight(20) }
            
            Write-Host "│ $id │ $type │ $title │ $state │ $assignedTo │" -ForegroundColor White
        }
        Write-Host "└──────┴──────────────┴────────────────────────────────┴─────────────┴──────────────────────┘" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Affichage des erreurs
    if ($Analysis.Errors.Count -gt 0) {
        Write-Host "❌ ERREURS DÉTECTÉES:" -ForegroundColor Red
        foreach ($err in $Analysis.Errors) {
            Write-Host "   • $err" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    # Affichage des avertissements
    if ($Analysis.Warnings.Count -gt 0) {
        Write-Host "⚠️  AVERTISSEMENTS:" -ForegroundColor Yellow
        foreach ($warning in $Analysis.Warnings) {
            Write-Host "   • $warning" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Résumé final
    Write-Host "✅ RÉSUMÉ:" -ForegroundColor Green
    Write-Host "   • Créations: $($Analysis.CreationItems.Count) work items" -ForegroundColor White
    Write-Host "   • Modifications: $($Analysis.UpdateItems.Count) work items" -ForegroundColor White
    Write-Host "   • Erreurs: $($Analysis.Errors.Count)" -ForegroundColor $(if ($Analysis.Errors.Count -gt 0) { "Red" } else { "White" })
    Write-Host "   • Avertissements: $($Analysis.Warnings.Count)" -ForegroundColor $(if ($Analysis.Warnings.Count -gt 0) { "Yellow" } else { "White" })
    Write-Host ""
}

# Fonction pour demander confirmation avant l'import
function Confirm-ImportExecution {
    param(
        [hashtable]$Analysis,
        [switch]$Force,
        [switch]$WhatIf
    )
    
    # En mode WhatIf, on affiche juste et on continue
    if ($WhatIf) {
        Write-Host "🔍 MODE SIMULATION ACTIVÉ - Aucune modification ne sera effectuée" -ForegroundColor Magenta
        Write-Host ""
        return $true
    }
    
    # En mode Force, on skip la confirmation
    if ($Force) {
        Write-Host "⚡ MODE FORCE ACTIVÉ - Import sans confirmation" -ForegroundColor Magenta
        Write-Host ""
        return $true
    }
    
    # Si des erreurs critiques, arrêt immédiat
    if ($Analysis.Errors.Count -gt 0) {
        Write-Host "🚫 IMPORT IMPOSSIBLE:" -ForegroundColor Red
        Write-Host "   Des erreurs critiques ont été détectées. Corrigez le fichier CSV avant de continuer." -ForegroundColor Red
        Write-Host ""
        return $false
    }
    
    # Si pas d'opérations à effectuer
    if ($Analysis.CreationItems.Count -eq 0 -and $Analysis.UpdateItems.Count -eq 0) {
        Write-Host "ℹ️  AUCUNE OPÉRATION À EFFECTUER" -ForegroundColor Gray
        Write-Host "   Le fichier CSV ne contient aucune modification à appliquer." -ForegroundColor Gray
        Write-Host ""
        return $false
    }
    
    # Demande de confirmation
    Write-Host "🤔 Voulez-vous continuer avec cet import ?" -ForegroundColor Cyan
    if ($Analysis.Warnings.Count -gt 0) {
        Write-Host "   (Des avertissements ont été détectés ci-dessus)" -ForegroundColor Yellow
    }
    
    do {
        $response = Read-Host "   Tapez 'y' pour continuer, 'n' pour annuler"
        $response = $response.ToLower().Trim()
    } while ($response -notin @('y', 'yes', 'n', 'no', ''))
    
    if ($response -in @('y', 'yes')) {
        Write-Host ""
        Write-Host "✅ Import confirmé par l'utilisateur" -ForegroundColor Green
        Write-Host ""
        return $true
    } else {
        Write-Host ""
        Write-Host "❌ Import annulé par l'utilisateur" -ForegroundColor Red
        Write-Host ""
        return $false
    }
}

# Nouvelle fonction pour récupérer les IDs par plages (méthode de secours)
function Get-WorkItemIdsByRange {
    param([string]$Project)
    
    Write-Log "Utilisation de la méthode de récupération par plages d'IDs..." "Info"
    
    $allIds = @()
    $rangeSize = 10000
    $startId = 1
    $maxId = 999999  # ID maximum raisonnable
    $consecutiveEmpty = 0
    $maxConsecutiveEmpty = 3  # Arrêter après 3 plages vides consécutives
    
    while ($startId -lt $maxId -and $consecutiveEmpty -lt $maxConsecutiveEmpty) {
        $endId = $startId + $rangeSize - 1
        
        Write-Log "  Recherche dans la plage d'IDs: $startId à $endId..." "Info"
        
        $rangeQuery = @{
            query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$Project' AND [System.Id] >= $startId AND [System.Id] <= $endId AND [System.WorkItemType] IN ('Epic', 'Feature', 'Product Backlog Item', 'Task', 'User Story', 'Bug') ORDER BY [System.Id]"
        } | ConvertTo-Json
        
        $wiqlUri = "$Global:BaseUri/wit/wiql?api-version=6.0"
        
        try {
            $result = Invoke-RestMethod -Uri $wiqlUri -Method Post -Body $rangeQuery -Headers $Global:Headers
            
            if ($result.workItems.Count -gt 0) {
                $ids = $result.workItems | ForEach-Object { $_.id }
                $allIds += $ids
                Write-Log "    Trouvé: $($ids.Count) work items (IDs: $($ids[0]) à $($ids[-1]))" "Info"
                $consecutiveEmpty = 0
            }
            else {
                $consecutiveEmpty++
                Write-Log "    Aucun work item trouvé dans cette plage" "Info"
            }
        }
        catch {
            Write-Log "    Erreur lors de la requête pour la plage $startId-$endId : $($_.Exception.Message)" "Warning"
            $consecutiveEmpty++
        }
        
        $startId = $endId + 1
        
        # Si on a déjà trouvé des items et qu'on a plusieurs plages vides, on peut s'arrêter
        if ($allIds.Count -gt 0 -and $consecutiveEmpty -ge 2) {
            Write-Log "  Arrêt de la recherche après $consecutiveEmpty plages vides consécutives" "Info"
            break
        }
    }
    
    Write-Log "Récupération par plages terminée: $($allIds.Count) IDs trouvés" "Success"
    return $allIds | Sort-Object -Unique
}

# Point d'entrée principal
function Main {
    Write-Log "Démarrage du script Azure DevOps Items Importer/Exporter" "Info"
    Write-Log "Organisation: $Organization" "Info"
    Write-Log "Projet: $Project" "Info"
    Write-Log "Action: $Action" "Info"
    
    # Test de la connexion
    if (-not (Test-AzureDevOpsConnection)) {
        Write-Log "Impossible de continuer sans connexion valide." "Error"
        exit 1
    }
    
    try {
        switch ($Action) {
            "Export" {
                Write-Log "Début de l'export vers: $CsvFilePath" "Info"
                Export-WorkItemsToCsv -FilePath $CsvFilePath
            }
            "Import" {
                Write-Log "Début de l'import depuis: $CsvFilePath" "Info"
                Import-WorkItemsFromCsv -FilePath $CsvFilePath
            }
        }
        
        Write-Log "Opération '$Action' terminée avec succès!" "Success"
    }
    catch {
        Write-Log "Échec de l'opération '$Action': $($_.Exception.Message)" "Error"
        exit 1
    }
}

# Exécution du script principal
Main
