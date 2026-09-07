$enJson = Get-Content .\project.json -Raw | ConvertFrom-Json
$frJson = Get-Content .\old_vf.json -Raw | ConvertFrom-Json

# --- Mapping id de set -> nom du set, pour chaque fichier ---
function Get-SetNameMap($json) {
    $map = @{}
    foreach ($set in $json.encounter_sets) { $map[$set.id] = $set.name }
    return $map
}
$enSetNames = Get-SetNameMap $enJson
$frSetNames = Get-SetNameMap $frJson

# --- Empreinte "illustration" d'une carte (front + back) ---
function Get-CardFingerprint($card) {
    $parts = @()
    foreach ($side in @('front','back')) {
        if ($card.$side) {
            foreach ($field in @('illustration','illustration_scale','illustration_pan_x','illustration_pan_y','illustration_rotation')) {
                if ($card.$side.PSObject.Properties.Name -contains $field) {
                    $parts += "$side.$field=$($card.$side.$field)"
                }
            }
        }
    }
    return ($parts -join '|')
}

# --- Index des cartes EN par empreinte (une empreinte peut avoir plusieurs exemplaires) ---
$enByFingerprint = @{}
foreach ($card in $enJson.cards) {
    $fp = Get-CardFingerprint $card
    if (-not $fp) { continue }
    if (-not $enByFingerprint.ContainsKey($fp)) {
        $enByFingerprint[$fp] = [System.Collections.Generic.List[object]]::new()
    }
    $enByFingerprint[$fp].Add($card)
}

$used = @{}   # id EN déjà attribués
$results = [System.Collections.Generic.List[object]]::new()
$stillUnmatched = [System.Collections.Generic.List[object]]::new()

foreach ($frCard in $frJson.cards) {
    $fp = Get-CardFingerprint $frCard
    $match = $null

    if ($fp -and $enByFingerprint.ContainsKey($fp)) {
        $setName = $frSetNames[$frCard.encounter_set]
        # priorité : même nom de set, pas encore utilisé
        $candidates = $enByFingerprint[$fp] | Where-Object { -not $used.ContainsKey($_.id) }
        $best = $candidates | Where-Object { $enSetNames[$_.encounter_set] -eq $setName } | Select-Object -First 1
        if (-not $best) { $best = $candidates | Select-Object -First 1 }
        if ($best) { $match = $best }
    }

    if ($match) {
        $used[$match.id] = $true
        $results.Add([PSCustomObject]@{
            fr_id = $frCard.id; fr_name = $frCard.name
            en_id = $match.id;  en_name = $match.name
            set   = $frSetNames[$frCard.encounter_set]
            methode = "illustration"
        })
    } else {
        $stillUnmatched.Add($frCard)
    }
}

# --- Repli pour les cartes sans illustration (ex: cartes chaos) ---
# Regroupe par (nom du set, type, niveau) et apparie par ordre d'encounter_number
function Get-FallbackKey($card, $setNameMap) {
    $type = $card.front.type
    $level = $card.front.level
    $setName = $setNameMap[$card.encounter_set]
    return "$setName|$type|$level"
}

$enRemaining = $enJson.cards | Where-Object { -not $used.ContainsKey($_.id) }
$enGroups = $enRemaining | Group-Object { Get-FallbackKey $_ $enSetNames }
$enGroupMap = @{}
foreach ($g in $enGroups) {
    $enGroupMap[$g.Name] = $g.Group | Sort-Object { [int]($_.encounter_number -replace '\D','') }
}

$frGroups = $stillUnmatched | Group-Object { Get-FallbackKey $_ $frSetNames }
$finalUnmatched = [System.Collections.Generic.List[object]]::new()

foreach ($g in $frGroups) {
    $frSorted = $g.Group | Sort-Object { [int]($_.encounter_number -replace '\D','') }
    $enCandidates = $enGroupMap[$g.Name]
    for ($i = 0; $i -lt $frSorted.Count; $i++) {
        if ($enCandidates -and $i -lt $enCandidates.Count) {
            $frCard = $frSorted[$i]
            $enCard = $enCandidates[$i]
            $used[$enCard.id] = $true
            $results.Add([PSCustomObject]@{
                fr_id = $frCard.id; fr_name = $frCard.name
                en_id = $enCard.id; en_name = $enCard.name
                set   = $frSetNames[$frCard.encounter_set]
                methode = "repli (set+type+niveau)"
            })
        } else {
            $finalUnmatched.Add($frSorted[$i])
        }
    }
}

# --- Export des résultats ---
$results | Export-Csv -Path .\fr_to_en_mapping.csv -NoTypeInformation -Encoding UTF8
$finalUnmatched | Select-Object id, name, encounter_number, encounter_set |
    Export-Csv -Path .\fr_unmatched.csv -NoTypeInformation -Encoding UTF8

Write-Host "Cartes appariées : $($results.Count) / $($frJson.cards.Count)"
Write-Host "Cartes non trouvées : $($finalUnmatched.Count) (voir fr_unmatched.csv)"