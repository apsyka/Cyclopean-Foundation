<#
.SYNOPSIS
    Ajoute le champ "collection_icon" sur front ET back de chaque carte chaos
    d'un project.json.

.DESCRIPTION
    Pour chaque carte dont front.type = "chaos" :
      - Si front.collection_icon n'existe pas -> on l'ajoute, valeur = icon
        du projet (champ "icon" a la racine du json).
      - Si front.collection_icon existe deja  -> on ne touche a rien, mais on
        affiche un avertissement (nom de la carte + valeur actuelle).
    Meme traitement, independamment, sur back.collection_icon.

.PARAMETER JsonPath
    Chemin vers project.json en entree.

.PARAMETER OutputJsonPath
    Chemin de sortie. Par defaut : le meme fichier avec un suffixe "_updated".

.NOTES
    A executer avec PowerShell 7+ (pwsh) de preference.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputJsonPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputJsonPath) {
    $dir = Split-Path -Parent $JsonPath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($JsonPath)
    $OutputJsonPath = Join-Path $dir "$name`_updated.json"
}

Write-Host "Lecture du json : $JsonPath"
$jsonRaw = Get-Content -Path $JsonPath -Raw -Encoding UTF8
$data = $jsonRaw | ConvertFrom-Json -Depth 100
Write-Host "  $($data.cards.Count) carte(s) chargees."

$projectIcon = "<project.icon>"
Write-Host "  Icone du projet : '$projectIcon'"
Write-Host ""

function Set-CollectionIcon {
    param(
        [Parameter(Mandatory = $true)] $Face,        # front ou back (objet)
        [Parameter(Mandatory = $true)] [string]$FaceName,  # "front" ou "back", pour les messages
        [Parameter(Mandatory = $true)] [string]$CardName,
        [Parameter(Mandatory = $true)] [string]$ProjectIcon
    )

    if ($null -eq $Face) { return $false }

    $hasField = $Face.PSObject.Properties.Name -contains 'collection_icon'

    if ($hasField) {
        Write-Warning "Carte '$CardName' : '$FaceName.collection_icon' existe deja (valeur actuelle : '$($Face.collection_icon)') -> inchange."
        return $false
    }
    else {
        $Face | Add-Member -MemberType NoteProperty -Name 'collection_icon' -Value $ProjectIcon -Force
        return $true
    }
}

$addedCount = 0
$warnedCount = 0
$chaosTotal = 0

foreach ($card in $data.cards) {
    if ($null -eq $card) { continue }
    if ($card.front.type -ne 'chaos') { continue }

    $chaosTotal++
    $cardName = $card.name

    $beforeAdded = $addedCount

    if (Set-CollectionIcon -Face $card.front -FaceName 'front' -CardName $cardName -ProjectIcon $projectIcon) {
        $addedCount++
    }
    else {
        $warnedCount++
    }

    if (Set-CollectionIcon -Face $card.back -FaceName 'back' -CardName $cardName -ProjectIcon $projectIcon) {
        $addedCount++
    }
    else {
        $warnedCount++
    }
}

Write-Host ""
Write-Host "Cartes chaos traitees : $chaosTotal"
Write-Host "Champs collection_icon ajoutes : $addedCount"
Write-Host "Champs deja presents (avertissements ci-dessus, inchanges) : $warnedCount"

Write-Host ""
Write-Host "Ecriture du json corrige : $OutputJsonPath"
$data | ConvertTo-Json -Depth 100 | Set-Content -Path $OutputJsonPath -Encoding UTF8

Write-Host "Termine."