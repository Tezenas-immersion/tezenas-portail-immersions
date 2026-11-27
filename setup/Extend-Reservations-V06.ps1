$ErrorActionPreference = "Stop"
$TenantID = "a7f315c0-79f6-4c05-92b6-c9396003e258"
$SiteID = "groupetezenas.sharepoint.com,9a59d1c6-6e28-4907-a1c5-c68922f8748d,fbc51b36-5df1-4f50-8c0e-bb49f8ff3ac7"
$Base = "https://graph.microsoft.com/v1.0/sites/$SiteID/lists"
Connect-MgGraph -TenantId $TenantID -Scopes "Sites.Manage.All" -ContextScope Process -NoWelcome
function Get-Columns($Uri) {
    do {
        $Page = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $Page.value
        $Uri = $Page.'@odata.nextLink'
    } while ($Uri)
}
$Schema = [ordered]@{
    Civilite = "Text"
    DateNaissance = "Text"
    VilleResidence = "Text"
    CodePostalResidence = "Text"
    Terminale = "Text"
    SpecialitesBac = "Text"
    AdresseEtablissement = "LongText"
    EmailEtablissement = "Text"
    TelephoneEtablissement = "Text"
    CiviliteChef = "Text"
    ChefEtablissement = "Text"
    Acceptation = "Boolean"
    DateAcceptation = "DateTime"
    VersionAcceptation = "Text"
    PlaceUnique = "Unique"
    DemandeUnique = "Unique"
}
try {
    $Uri = "$Base/9da020ad-1265-43bd-b611-6e8c209908a4/columns"
    $Columns = @(Get-Columns $Uri)
    foreach ($Name in $Schema.Keys) {
        $Existing = $Columns | Where-Object { $_.name -eq $Name } | Select-Object -First 1
        if ($Existing) { Write-Host "[EXISTE] $Name"; continue }
        $Definition = @{ name = $Name; displayName = $Name }
        switch ($Schema[$Name]) {
            "Text" { $Definition.text = @{} }
            "LongText" { $Definition.text = @{ allowMultipleLines = $true } }
            "Boolean" { $Definition.boolean = @{} }
            "DateTime" { $Definition.dateTime = @{ format = "dateTime" } }
            "Unique" {
                $Definition.text = @{}
                $Definition.indexed = $true
                $Definition.enforceUniqueValues = $true
            }
        }
        Invoke-MgGraphRequest -Method POST -Uri $Uri -Body ($Definition | ConvertTo-Json -Depth 8) -ContentType "application/json" | Out-Null
        Write-Host "[OK] $Name" -ForegroundColor Green
    }
    $Columns = @(Get-Columns $Uri)
    $Missing = @($Schema.Keys | Where-Object { $Columns.name -notcontains $_ })
    if ($Missing.Count) { throw "Colonnes manquantes : $($Missing -join ', ')" }
    foreach ($Name in @("PlaceUnique","DemandeUnique")) {
        $C = $Columns | Where-Object { $_.name -eq $Name }
        if ($C.enforceUniqueValues -ne $true -or $C.indexed -ne $true) { throw "Contrainte non validee : $Name" }
        Write-Host "[UNICITE OK] $Name" -ForegroundColor Green
    }
    Write-Host "SCHEMA V0.6 OK - aucune reservation creee." -ForegroundColor Green
} finally { Disconnect-MgGraph | Out-Null }
