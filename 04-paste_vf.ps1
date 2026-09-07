# --- Charger les deux fichiers ---
$projectFr = Get-Content .\project_fr.json     -Raw | ConvertFrom-Json
$oldVfUpdated = Get-Content .\old_vf_updated.json -Raw | ConvertFrom-Json

# --- Index des cartes old_vf_updated par id (cards = LISTE avec champ .id) ---
$oldVfById = @{}
foreach ($card in $oldVfUpdated.cards) {
    if ($card.id) {
        $oldVfById[$card.id] = $card
    }
}
Write-Host "Cartes chargées depuis old_vf_updated : $($oldVfById.Count)"

# --- Fonction de fusion : ne remplace QUE les champs déjà présents dans $target ---
function Merge-ExistingFields {
    param($target, $source)

    foreach ($prop in $target.PSObject.Properties) {
        $name = $prop.Name

        if ($source.PSObject.Properties.Name -notcontains $name) {
            # le champ n'existe pas dans la source -> on laisse tel quel
            continue
        }

        $sourceValue = $source.$name

        if ($prop.Value -is [System.Management.Automation.PSCustomObject] -and
            $sourceValue -is [System.Management.Automation.PSCustomObject]) {
            # objet imbriqué (ex: front, back) -> fusion récursive
            Merge-ExistingFields -target $prop.Value -source $sourceValue
        }
        else {
            # valeur simple, tableau, etc. -> remplacement direct
            $target.$name = $sourceValue
        }
    }
}

# --- project_fr.json : cards est un OBJET indexé par id (guid = clé) ---
# IMPORTANT : forcer en tableau réel avec @(...) pour que .Count soit fiable
$cardEntries = @($projectFr.cards.PSObject.Properties)
$totalCards = $cardEntries.Count

$updated = 0
$notFound = [System.Collections.Generic.List[object]]::new()

foreach ($entry in $cardEntries) {
    $id = $entry.Name          # le guid
    $frCard = $entry.Value     # l'objet { name, front, back }

    if ($oldVfById.ContainsKey($id)) {
        $sourceCard = $oldVfById[$id]
        Merge-ExistingFields -target $frCard -source $sourceCard
        $updated++
    }
    else {
        $notFound.Add([PSCustomObject]@{
                id   = $id
                name = $frCard.name
            })
    }
}

Write-Host "Cartes mises a jour : $updated / $totalCards"

if ($notFound.Count -gt 0) {
    Write-Host "Cartes de project_fr.json sans correspondance dans old_vf_updated : $($notFound.Count)"
    $notFound | Export-Csv -Path .\project_fr_ids_non_trouves.csv -NoTypeInformation -Encoding UTF8
}

# --- Sauvegarder le resultat ---
$projectFr | ConvertTo-Json -Depth 20 | Out-File -FilePath .\project_fr_final.json -Encoding utf8

Write-Host "Fichier genere : project_fr_final.json"