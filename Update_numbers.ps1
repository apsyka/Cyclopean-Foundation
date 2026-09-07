<#
.SYNOPSIS
    Reporte les numeros (Card Number / Encounter Set Number) du xlsx vers
    project_number / encounter_number dans project.json.

.DESCRIPTION
    - Card Number (xlsx)          -> project_number (json), en entier.
    - Encounter Set Number (xlsx) -> encounter_number (json), MAIS sans le
      total "/z" : "19-20/27" devient "19-20", "5/27" devient "5".

    L'appariement entre une ligne du xlsx et une carte du json se fait par :
      (nom du set d'encounter, nom de carte normalise, type de face avant)
    et non par un identifiant, car les UUID des deux fichiers ne correspondent
    pas entre eux (deux exports differents du meme projet).

    Les cartes "Agenda"/"Act" du xlsx portent un prefixe ("Agenda 1a ", "Act 2a ")
    absent du nom dans le json : ce prefixe est retire avant comparaison.
    Les cartes doubles (front "story" qui bascule en "enemy"/"asset" au dos)
    sont aussi prises en compte via une liste de types de repli.

.PARAMETER ExcelPath
    Chemin vers le fichier .xlsx corrige (Cyclopean_Foundations_corrige.xlsx).

.PARAMETER JsonPath
    Chemin vers project.json en entree.

.PARAMETER OutputJsonPath
    Chemin de sortie. Par defaut : le meme fichier avec un suffixe "_updated".

.NOTES
    Necessite le module ImportExcel (Install-Module ImportExcel) pour lire
    le .xlsx sans avoir Excel installe.
    A executer avec PowerShell 7+ (pwsh) de preference.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputJsonPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputJsonPath) {
    $dir  = Split-Path -Parent $JsonPath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($JsonPath)
    $OutputJsonPath = Join-Path $dir "$name`_updated.json"
}

# --- Verifie / charge le module ImportExcel -------------------------------
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Warning "Le module 'ImportExcel' n'est pas installe."
    Write-Warning "Installez-le avec : Install-Module ImportExcel -Scope CurrentUser"
    throw "Module ImportExcel manquant."
}
Import-Module ImportExcel

# --- Charge les donnees -----------------------------------------------------
Write-Host "Lecture du xlsx : $ExcelPath"
$rows = Import-Excel -Path $ExcelPath

Write-Host "Lecture du json : $JsonPath"
$jsonRaw = Get-Content -Path $JsonPath -Raw -Encoding UTF8
$data    = $jsonRaw | ConvertFrom-Json -Depth 100

# --- Table de correspondance nom de set -> id de set -----------------------
$setIdByName = @{}
foreach ($es in $data.encounter_sets) {
    $setIdByName[$es.name] = $es.id
}

function Get-PlainSetName {
    param([string]$XlsxSetName)
    # "01 - Lost Moorings" -> "Lost Moorings" ; "Grand Compass" -> "Grand Compass"
    if ($XlsxSetName -match '^\s*\d+\s*-\s*(.+)$') {
        return $Matches[1].Trim()
    }
    return $XlsxSetName.Trim()
}

function Get-NormalizedName {
    param([string]$CardName)
    # Retire le prefixe "Agenda 1a "/"Act 2b " etc. puis ne garde que alphanum, en minuscule.
    $s = $CardName -replace '^(Agenda|Act)\s+\d+[a-zA-Z]?\s+', ''
    $s = ($s -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    return $s
}

# xlsx Type -> liste ordonnee des front.type a essayer dans le json
$fallbackTypes = @{
    'enemy'     = @('enemy', 'story')
    'asset'     = @('asset', 'story')
    'treachery' = @('treachery', 'story')
    'agenda'    = @('agenda')
    'act'       = @('act')
    'location'  = @('location')
    'chaos'     = @('chaos')
    'story'     = @('story')
}

# --- Construit les "pools" de cartes candidates par cle --------------------
# Cle = "<set_id>|<nom_normalise>|<front_type>" -> file (FIFO) de cartes candidates
$pools = @{}
foreach ($card in $data.cards) {
    $ftype = $card.front.type
    $key = "{0}|{1}|{2}" -f $card.encounter_set, (Get-NormalizedName $card.name), $ftype
    if (-not $pools.ContainsKey($key)) {
        $pools[$key] = [System.Collections.Generic.Queue[object]]::new()
    }
    $pools[$key].Enqueue($card)
}

# --- Applique la correspondance, ligne par ligne du xlsx (ordre Card Number) -
$rows = $rows | Sort-Object { [int]$_.'Card Number' }

$matchedCount   = 0
$unmatchedRows  = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $setId  = $setIdByName[(Get-PlainSetName $row.'Encounter Set Name')]
    $nname  = Get-NormalizedName $row.'English Card Name'
    $xtype  = $row.Type
    $tries  = $fallbackTypes[$xtype]
    if (-not $tries) { $tries = @($xtype) }

    $found = $null
    foreach ($ftype in $tries) {
        $key = "{0}|{1}|{2}" -f $setId, $nname, $ftype
        if ($pools.ContainsKey($key) -and $pools[$key].Count -gt 0) {
            $found = $pools[$key].Dequeue()
            break
        }
    }

    if ($found) {
        # Card Number -> project_number (entier)
        $found.project_number = [int]$row.'Card Number'

        # Encounter Set Number -> encounter_number, SANS le total "/z"
        $encSetNumber = [string]$row.'Encounter Set Number'
        $found.encounter_number = ($encSetNumber -replace '/.*$', '')

        $matchedCount++
    }
    else {
        $unmatchedRows.Add($row) | Out-Null
    }
}

# --- Rapport -----------------------------------------------------------------
Write-Host ""
Write-Host "Cartes mises a jour : $matchedCount / $($rows.Count)"

if ($unmatchedRows.Count -gt 0) {
    Write-Warning "Cartes NON trouvees dans le json (a verifier manuellement) :"
    foreach ($u in $unmatchedRows) {
        Write-Warning ("  - [{0}] {1} (type: {2})" -f $u.'Encounter Set Name', $u.'English Card Name', $u.Type)
    }
}

# --- Sauvegarde ---------------------------------------------------------------
Write-Host ""
Write-Host "Ecriture du json corrige : $OutputJsonPath"
$data | ConvertTo-Json -Depth 100 | Set-Content -Path $OutputJsonPath -Encoding UTF8

Write-Host "Termine."