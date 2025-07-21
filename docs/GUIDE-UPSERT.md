# Guide d'utilisation de la fonctionnalité Upsert

## Vue d'ensemble
La fonctionnalité **upsert** (update + insert) permet de :
- **Créer** de nouveaux work items quand l'ID est vide, 0, ou n'existe pas
- **Mettre à jour** des work items existants quand l'ID existe dans Azure DevOps

## Format CSV pour l'upsert

### Pour CRÉER un nouveau work item :
- Laissez le champ `Id` **vide** ou mettez `0`
- Remplissez tous les champs obligatoires : `Type`, `Title`, `State`

```csv
Id,Type,Title,State,AssignedTo,Description,RemainingWork,Tags,IterationPath,AreaPath
,Task,Nouvelle tâche,New,user@company.com,Description de la tâche,8,nouveau;tâche,Sprint 1,Backend
0,Bug,Nouveau bug,New,dev@company.com,Bug description,4,bug;nouveau,Sprint 1,Frontend
```

### Pour METTRE À JOUR un work item existant :
- Spécifiez l'`Id` du work item existant
- Remplissez les champs à modifier

```csv
Id,Type,Title,State,AssignedTo,Description,RemainingWork,Tags,IterationPath,AreaPath
1001,Epic,Epic modifié,Active,manager@company.com,Epic mis à jour,,epic;modifié,Release 1,Infrastructure
```

## Exemples d'utilisation

### 1. Test en mode simulation (recommandé)
```powershell
.\AzureDevOpsItemsImporter.ps1 -Organization "votre-org" -Project "votre-projet" -Action Import -CsvFilePath "exemple-import-upsert.csv" -WhatIf
```

### 2. Import réel après validation
```powershell
.\AzureDevOpsItemsImporter.ps1 -Organization "votre-org" -Project "votre-projet" -Action Import -CsvFilePath "exemple-import-upsert.csv"
```

### 3. Utilisation avec la configuration
```powershell
# Après avoir exécuté Setup.ps1
. .\config.ps1
.\AzureDevOpsItemsImporter.ps1 -Organization $Organization -Project $Project -Action Import -CsvFilePath "votre-fichier.csv" -WhatIf
```

## Champs obligatoires

### Pour CRÉATION :
- `Type` : Task, Bug, User Story, Epic, etc.
- `Title` : Titre du work item
- `State` : New, Active, Resolved, Closed, etc.

### Pour MISE À JOUR :
- `Id` : ID numérique du work item existant
- Au moins un champ à modifier

## Champs optionnels
- `AssignedTo` : Email de l'assigné
- `Description` : Description détaillée
- `RemainingWork` : Charge restante (en heures)
- `Tags` : Tags séparés par des points-virgules
- `IterationPath` : Chemin de l'itération (ex: Projet\Sprint 1)
- `AreaPath` : Chemin de la zone (ex: Projet\Backend)

## Validation et sécurité

### 1. Validation préalable
```powershell
.\ValidateExport.ps1 -CsvFile "votre-fichier.csv"
```

### 2. Mode simulation obligatoire
- Toujours tester avec `-WhatIf` avant l'import réel
- Vérifiez les logs pour vous assurer que les bonnes actions seront effectuées

### 3. Logs détaillés
Le script affiche :
- ✅ Créations : `[SIMULATION] Création d'un work item de type X avec N champs`
- ✅ Mises à jour : `[SIMULATION] Mise à jour du work item ID avec N champs`
- ✅ Résumé : `Créés: X, Mis à jour: Y, Erreurs: Z`

## Messages d'erreur courants

### "Work item X n'existe pas - Création d'un nouveau work item"
- **Normal** : L'ID spécifié n'existe pas, le script va créer un nouveau work item

### "Aucun champ fourni pour créer le work item"
- **Solution** : Vérifiez que les champs obligatoires sont remplis (Type, Title, State)

### "Erreur lors de la création/mise à jour"
- **Solution** : Vérifiez les permissions Azure DevOps et la validité des champs

## Bonnes pratiques

1. **Testez toujours** avec `-WhatIf` avant l'import réel
2. **Validez le CSV** avec `ValidateExport.ps1` avant l'import
3. **Sauvegardez** votre projet avant les imports importants
4. **Utilisez des IDs** existants pour les mises à jour, laissez vide pour les créations
5. **Respectez les types** de work items autorisés dans votre projet Azure DevOps

## Dépannage

### Le script ne trouve pas les work items à mettre à jour
- Vérifiez que les IDs existent dans Azure DevOps
- Vérifiez les permissions de lecture sur le projet

### Les créations échouent
- Vérifiez les champs obligatoires pour le type de work item
- Vérifiez les permissions de création dans Azure DevOps
- Vérifiez que les valeurs de `State` sont valides pour le type de work item
