# Vygeneruje seznam RIF uzivatelu do src/autorun.js (blok mezi // USERS-BEGIN a // USERS-END).
# Spustit z korene slozky add-inu po Connect-ExchangeOnline do tenantu allriskreality.
#
# RIF_USERS  = kazda adresa @rifas.cz (primarni i alias) -> udaje do podpisu.
#              Sdilena schranka dostane jen { shared = true } - podpis se pak bere od toho, kdo pise.
# RIF_PEOPLE = primarni adresa uzivatele -> jeho udaje (pro psani ze sdilene schranky).

$domain = 'rifas.cz'

# Adresy, kterym add-in NIKDY nevlozi podpis (ponecha jim vlastni / zadny).
# Pr.: $exclude = @('noreply@rifas.cz')
$exclude = @(
) | ForEach-Object { $_.ToLower() }

# Schranky typu UserMailbox, ktere se maji chovat jako sdilene (napr. licencovana "dluhopisy@").
# Schranky bez jmena i prijmeni se jako sdilene berou automaticky.
$treatAsShared = @(
) | ForEach-Object { $_.ToLower() }

$jsPath  = Join-Path (Join-Path $PSScriptRoot 'src') 'autorun.js'
$users   = [ordered]@{}
$people  = [ordered]@{}
$pattern = "^smtp:[^@]+@$([regex]::Escape($domain))$"

$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited |
    Where-Object { $_.EmailAddresses -match $pattern }

foreach ($mb in $mailboxes) {
    $isShared = $mb.RecipientTypeDetails -eq 'SharedMailbox'
    $primary  = $mb.PrimarySmtpAddress.ToString().ToLower()
    if (-not $isShared) {
        $u = Get-User -Identity $mb.Identity
        # Schranka bez jmena a prijmeni = "spolecna" schranka vedena jako UserMailbox -> chovat se jako ke sdilene
        if ($treatAsShared -contains $primary -or (-not $u.FirstName -and -not $u.LastName)) {
            Write-Warning "${primary}: $($mb.RecipientTypeDetails) bez jmena -> beru jako sdilenou schranku"
            $isShared = $true
        }
    }
    $addrs = $mb.EmailAddresses |
        Where-Object { $_ -match $pattern } |
        ForEach-Object { ($_ -replace '^smtp:', '').ToLower() }

    if ($isShared) {
        $data = [ordered]@{ shared = $true }
    } else {
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
        $people[$primary] = $data
        if (-not $u.FirstName -or -not $u.LastName) { Write-Warning "$($mb.PrimarySmtpAddress): chybi FirstName/LastName" }
        if (-not $mb.CustomAttribute15)             { Write-Warning "$($mb.PrimarySmtpAddress): chybi CustomAttribute15 (fotka)" }
    }

    foreach ($a in $addrs) {
        if ($exclude -contains $a) { Write-Host "Vynechano (vyjimka): $a"; continue }
        $users[$a] = $data
    }
}

function To-Js($obj) {
    $j = $obj | ConvertTo-Json -Depth 3
    if (-not $j) { $j = '{}' }
    # Diakritiku zapsat jako \uXXXX, aby nezalezelo na kodovani souboru ani verzi PowerShellu
    [regex]::Replace($j, '[^\x00-\x7F]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
}

$block = "// USERS-BEGIN  (generuje export-users.ps1 - rucne needitovat)`r`n" +
         "var RIF_USERS = $(To-Js $users);`r`n" +
         "var RIF_PEOPLE = $(To-Js $people);`r`n" +
         "// USERS-END"
$js = [IO.File]::ReadAllText($jsPath)
$js = [regex]::Replace($js, '(?s)// USERS-BEGIN.*?// USERS-END',
          [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block })
[IO.File]::WriteAllText($jsPath, $js, (New-Object System.Text.UTF8Encoding $false))

Write-Host "Zapsano $($users.Count) adres @$domain a $($people.Count) osob do $jsPath"
