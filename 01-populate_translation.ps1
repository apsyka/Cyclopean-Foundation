$json = Get-Content .\project.json -Raw | ConvertFrom-Json
$cards = $json.cards
$field_list = @('name', 'text', 'flavor_text', 'subtitle', 'victory', 'traits', 'entries', 'difficulty', 'chaos_extra', 'tracking')
$exported = [ordered]@{}

foreach ($card in $cards) {
    $id = $card.id

    $cardEntry = [ordered]@{
        name = $card.name
    }

    foreach ($side in @('front', 'back')) {
        if ($card.PSObject.Properties.Name -contains $side) {
            $sideObj = [ordered]@{}
            foreach ($field in $card.$side.psobject.Properties.Name) {
                if ($field_list -contains $field) {
                    $sideObj[$field] = $card.$side.$field
                }
            }
            if ($sideObj.Count -gt 0) {
                $cardEntry[$side] = $sideObj
            }
        }
    }

    $exported[$id] = $cardEntry
}
$exported | ConvertTo-Json -Depth 10 | Out-File -FilePath .\exported.json -Encoding utf8