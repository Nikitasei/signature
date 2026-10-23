# Vygeneruje seznam RIF uzivatelu do src/autorun.js (blok mezi // USERS-BEGIN a // USERS-END).
# Spustit z korene slozky add-inu po Connect-ExchangeOnline do tenantu allriskreality.
# Klicem je KAZDA adresa @rifas.cz (primarni i alias), takze funguje i psani "jako RIF" z aliasu.

$domain   = 'rifas.cz'
$jsPath   = Join-Path $PSScriptRoot 'src\autorun.js'
$users    = [ordered]@{}

$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited |
    Where-Object { $_.EmailAddresses -match "^smtp:[^@]+@$([regex]::Escape($domain))$" }

foreach ($mb in $mailboxes) {
    $u = Get-User -Identity $mb.Identity
    $data = [ordered]@{
        firstName  = "$($u.FirstName)"
        lastName   = "$($u.LastName)"
        title      = "$($u.Title)"
        phone      = "$($u.Phone)"
        street     = "$($u.StreetAddress)"
        postalCode = "$($u.PostalCode)"
        city       = "$($u.City)"
        photo      = "$($mb.CustomAttribute15)"
    }
    $addrs = $mb.EmailAddresses |
        Where-Object { $_ -match "^smtp:[^@]+@$([regex]::Escape($domain))$" } |
        ForEach-Object { ($_ -replace '^smtp:', '').ToLower() }
    foreach ($a in $addrs) { $users[$a] = $data }

    if (-not $u.FirstName -or -not $u.LastName) { Write-Warning "$($mb.PrimarySmtpAddress): chybi FirstName/LastName" }
    if (-not $mb.CustomAttribute15)             { Write-Warning "$($mb.PrimarySmtpAddress): chybi CustomAttribute15 (fotka)" }
}

$json = $users | ConvertTo-Json -Depth 3
if (-not $json) { $json = '{}' }
# Diakritiku zapsat jako \uXXXX, aby nezalezelo na kodovani souboru ani verzi PowerShellu
$json = [regex]::Replace($json, '[^\x00-\x7F]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })

$block = "// USERS-BEGIN  (generuje export-users.ps1 - rucne needitovat)`r`nvar RIF_USERS = $json;`r`n// USERS-END"
$js    = [IO.File]::ReadAllText($jsPath)
$js    = [regex]::Replace($js, '(?s)// USERS-BEGIN.*?// USERS-END',
            [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block })
[IO.File]::WriteAllText($jsPath, $js, (New-Object System.Text.UTF8Encoding $false))

Write-Host "Zapsano $($users.Count) adres do $jsPath"
