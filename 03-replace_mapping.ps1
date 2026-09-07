# --- Charger le mapping fr_id -> en_id ---
$mapping = Import-Csv .\fr_to_en_mapping.csv

$idMap = @{}
foreach ($row in $mapping) {
    $idMap[$row.fr_id] = $row.en_id
}

Write-Host "Mapping chargé : $($idMap.Count) correspondances"

# --- Charger le fichier à corriger ---
$oldVf = Get-Content .\old_vf.json -Raw | ConvertFrom-Json

$replaced = 0
$notFound = [System.Collections.Generic.List[object]]::new()

foreach ($card in $oldVf.cards) {
    if ($idMap.ContainsKey($card.id)) {
        $card.id = $idMap[$card.id]
        $replaced++
    }
    else {
        $notFound.Add([PSCustomObject]@{
                id   = $card.id
                name = $card.name
            })
    }
}

Write-Host "IDs remplacés : $replaced / $($oldVf.cards.Count)"

if ($notFound.Count -gt 0) {
    Write-Host "IDs non trouvés dans le mapping : $($notFound.Count)"
    $notFound | Export-Csv -Path .\old_vf_ids_non_trouves.csv -NoTypeInformation -Encoding UTF8
}

# --- Sauvegarder le résultat ---
$oldVf | ConvertTo-Json -Depth 20 | Out-File -FilePath .\old_vf_updated.json -Encoding utf8

Write-Host "Fichier généré : old_vf_updated.json"