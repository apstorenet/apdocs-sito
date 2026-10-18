<#
apDocs - installazione, aggiornamento e disinstallazione su Windows.

  Installa per tutti gli utenti (serve amministratore, lo chiede da solo):
    irm https://apdocs.antoniopiccolo.com/installa.ps1 | iex

  Con opzioni (da file o con scriptblock):
    & ([scriptblock]::Create((irm https://apdocs.antoniopiccolo.com/installa.ps1))) -Utente
    .\installa.ps1 -Aggiorna | -Disinstalla | -RimuoviAltri | -Utente | -Silenzioso

apDocs e' basato su PdfCraft (storytold/pdfcraft, MIT OR Apache-2.0) di ArtCraft Team e contributori.
Lo script scarica la build ufficiale indicata in versione.json, ne verifica SHA256 e firma digitale
e la registra come programma PDF con nome e icone apDocs. Compatibile con Windows PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [switch]$Utente,          # solo per l'utente corrente (niente amministratore)
    [switch]$Aggiorna,        # controlla e applica una versione nuova (usato dall'attivita' pianificata)
    [switch]$Disinstalla,
    [switch]$RimuoviAltri,    # disinstalla gli altri programmi PDF (tiene firma digitale e Print to PDF)
    [switch]$Silenzioso,      # nessuna domanda, nessuna finestra
    [switch]$Attendi,
    [switch]$DopoScript,      # interno: -Aggiorna rilanciato dallo script appena aggiornato         # a fine lavoro aspetta Invio (usato quando lo script si rilancia come amministratore)
    [string]$Sorgente = 'https://apdocs.antoniopiccolo.com'   # URL o cartella locale con versione.json, brand/, installa.ps1
)

$SorgenteEsplicita = $PSBoundParameters.ContainsKey('Sorgente')
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$Nome = 'apDocs'
# Nome senza ".PDF" finale: su Windows 11 il servizio UCPD blocca PowerShell sulle chiavi che finiscono cosi'.
$ProgId = 'apDocs.Documento'
$NomeAttivita = 'apDocs - aggiornamento'

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal $id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- dove installare: tutti gli utenti (HKLM, Programmi) o utente corrente (HKCU, profilo) ---
# Un'installazione esistente decide da sola: per -Aggiorna e -Disinstalla si cerca prima quella per tutti.
if (-not $Utente -and ($Aggiorna -or $Disinstalla) -and -not (Test-Path "$env:ProgramFiles\$Nome\versione.txt") -and (Test-Path "$env:LOCALAPPDATA\Programs\$Nome\versione.txt")) {
    $Utente = $true
}
if ($Utente) {
    $Dir = "$env:LOCALAPPDATA\Programs\$Nome"
    $Hive = 'HKCU:'
    $StartMenu = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
} else {
    $Dir = "$env:ProgramFiles\$Nome"
    $Hive = 'HKLM:'
    $StartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
}
$Classes = "$Hive\Software\Classes"
$AppDir = Join-Path $Dir 'app'
$Exe = Join-Path $AppDir 'pdfcraft.exe'
$Registro = Join-Path $Dir 'registro.log'

function Scrivi([string]$testo) {
    $riga = '{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $testo
    if (-not $Silenzioso) { Write-Host $testo }
    try {
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
        Add-Content -Path $Registro -Value $riga -Encoding UTF8
    } catch { }
}

# Meta' pubblica della chiave di firma apDocs. Ogni versione.json scaricato deve essere firmato con la
# chiave privata corrispondente (che resta a apStore.net): dentro ci sono le impronte SHA256 di questo
# script e delle icone, quindi nulla di cio' che arriva dal sito viene usato senza quella firma.
$ChiavePubblica = '<RSAKeyValue><Modulus>0U0rR+uPYM1FVb6ZbuZ2Y7V81pXBhrPVd5Tj5fasrdp8iqbGTMgNl+cQ8bbvoNPiQ8+u6EELVqI7sZhkN7/7BGlGYGPmqzpDTWpJPhTCqpAKGbK5cvJPNbL5AIhLf444UGHScq5ve+YRFccSZq39FR6sou8A97cZzsYJMKlL51i8eyEmIYGyq+RYpCTR00SKHWBsdfyXueQaJOOjWMVp+drUbz2UXqYVyKbr9TWye3csFQAr2hb78l2LmrdFG1zLrj5NS87IPDU5mV0hXOxiPxLR22104iWgWzv/49JHZcBSm5qDk8jJyzfo0EO2J2Wj7S8JOTUX8X7RvTm2s+XH43UG0urT9MVPfwLbMk4l8l8368UCuQsCPuW32YzHBpLQc/VIHoNzNtYN9xTJAYpCvHv5Nyhor4OWnhvpjyzruEAtLJSH5hdlkIAKszfVHiB0QToBTtP+nmjdi6W7ukPkyCns8cg/WkWsEIdGK9CrW8YhSDtutj7NspU8xFQS2/QZ</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'

function Leggi-Byte([string]$percorso) {
    $src = $Sorgente.TrimEnd('/', '\')
    if ($src -match '^https?://') {
        $tmp = [IO.Path]::GetTempFileName()
        try {
            Invoke-WebRequest -Uri "$src/$($percorso -replace '\\', '/')" -OutFile $tmp -UseBasicParsing
            return [IO.File]::ReadAllBytes($tmp)
        } finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
    return [IO.File]::ReadAllBytes((Join-Path $src $percorso))
}

function Verifica-Firma([byte[]]$dati, [string]$firmaB64) {
    $rsa = New-Object Security.Cryptography.RSACryptoServiceProvider
    $rsa.PersistKeyInCsp = $false
    $rsa.FromXmlString($ChiavePubblica)
    try { return $rsa.VerifyData($dati, 'SHA256', [Convert]::FromBase64String($firmaB64.Trim())) }
    catch { return $false }
    finally { $rsa.Dispose() }
}

function Sha256Byte([byte[]]$dati) {
    $h = [Security.Cryptography.SHA256]::Create()
    try { return (-join ($h.ComputeHash($dati) | ForEach-Object { $_.ToString('x2') })) } finally { $h.Dispose() }
}

$script:Manifesto = $null
function Leggi-Manifesto {
    $dati = Leggi-Byte 'versione.json'
    $firma = [Text.Encoding]::ASCII.GetString((Leggi-Byte 'versione.json.sig'))
    if (-not (Verifica-Firma $dati $firma)) { throw 'Firma di versione.json non valida: file non pubblicato da apStore.net. Mi fermo.' }
    $testo = [Text.Encoding]::UTF8.GetString($dati).TrimStart([char]0xFEFF)
    $script:Manifesto = $testo | ConvertFrom-Json
    return $script:Manifesto
}

# Scarica un file del sito e lo scrive solo se la sua impronta e' quella del manifesto firmato.
function Scarica-Verificato([string]$percorso, [string]$destinazione) {
    $atteso = $script:Manifesto.file.($percorso -replace '\\', '/')
    if (-not $atteso) { throw "Il manifesto firmato non elenca $percorso." }
    $dati = Leggi-Byte $percorso
    if ((Sha256Byte $dati) -ne $atteso.ToLower()) { throw "Impronta di $percorso diversa da quella firmata. Mi fermo." }
    $tmp = "$destinazione.nuovo"
    [IO.File]::WriteAllBytes($tmp, $dati)
    Move-Item $tmp $destinazione -Force
}

function Architettura {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { return 'arm64' }
    return 'x64'
}

function Versione-Installata {
    $f = Join-Path $Dir 'versione.txt'
    if (Test-Path $f) { return (Get-Content $f -Raw).Trim() }
    return $null
}

# Scarica il pacchetto, verifica SHA256 e firma, estrae in una cartella temporanea e la restituisce.
function Prepara-Pacchetto($man) {
    $arch = Architettura
    $pac = $man.pacchetti.$arch
    if (-not $pac) { throw "Nessun pacchetto per l'architettura $arch nel manifesto." }
    $tmp = Join-Path $env:TEMP ("apdocs-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    $zip = Join-Path $tmp 'pacchetto.zip'
    Scrivi "Scarico $Nome $($man.versione) ($arch)..."
    Invoke-WebRequest -Uri $pac.url -OutFile $zip -UseBasicParsing
    $hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
    if ($hash -ne $pac.sha256.ToLower()) { Remove-Item $tmp -Recurse -Force; throw "Checksum errato: atteso $($pac.sha256), ottenuto $hash. Installazione annullata." }
    Scrivi "Checksum verificato."
    $estratto = Join-Path $tmp 'estratto'
    Expand-Archive -Path $zip -DestinationPath $estratto -Force
    $exe = Get-ChildItem $estratto -Recurse -Filter 'pdfcraft.exe' | Select-Object -First 1
    if (-not $exe) { Remove-Item $tmp -Recurse -Force; throw 'pdfcraft.exe non trovato nel pacchetto.' }
    # Le build ufficiali PdfCraft hanno la firma Authenticode del loro autore; le build apDocs no
    # (il manifesto non indica un firmatario) e valgono per l'impronta SHA256 nel manifesto firmato da noi.
    if ($man.firmatario) {
        $firma = Get-AuthenticodeSignature $exe.FullName
        if ($firma.Status -ne 'Valid' -or $firma.SignerCertificate.Subject -notmatch [regex]::Escape("CN=$($man.firmatario)")) {
            Remove-Item $tmp -Recurse -Force
            throw "Firma digitale non valida o di un altro firmatario ($($firma.Status), $($firma.SignerCertificate.Subject))."
        }
        Scrivi "Firma digitale verificata: $($man.firmatario)."
    }
    $radice = $exe.Directory.FullName
    # In Programmi gli utenti non possono scrivere: impostazioni e registri vanno nel profilo di ciascuno.
    Remove-Item (Join-Path $radice 'portable.txt') -Force -ErrorAction SilentlyContinue
    return @{ Tmp = $tmp; Radice = $radice }
}

function Applica-Pacchetto($pacchetto, [string]$versione) {
    $nuova = "$AppDir.nuova"
    $vecchia = "$AppDir.vecchia"
    Remove-Item $nuova, $vecchia -Recurse -Force -ErrorAction SilentlyContinue
    Copy-Item $pacchetto.Radice $nuova -Recurse
    if (Test-Path $AppDir) { Rename-Item $AppDir $vecchia }
    Rename-Item $nuova $AppDir
    Remove-Item $vecchia -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $pacchetto.Tmp -Recurse -Force -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $Dir 'versione.txt') -Value $versione -Encoding ASCII
}

function Copia-Script {
    $dest = Join-Path $Dir 'installa.ps1'
    if ($PSCommandPath -and (Test-Path $PSCommandPath)) {
        if ((Resolve-Path $PSCommandPath).Path -ne $dest) { Copy-Item $PSCommandPath $dest -Force }
    } else {
        Scarica-Verificato 'installa.ps1' $dest
    }
    Set-Content -Path (Join-Path $Dir 'sorgente.txt') -Value $Sorgente -Encoding ASCII
}

function Imposta([string]$chiave, [string]$nome, $valore) {
    try {
        if (-not (Test-Path $chiave)) { New-Item -Path $chiave -Force | Out-Null }
        if ($nome -eq '(default)') { Set-Item -Path $chiave -Value $valore }
        else { New-ItemProperty -Path $chiave -Name $nome -Value $valore -PropertyType String -Force | Out-Null }
    } catch [Security.SecurityException] {
        # Su Windows 11 il servizio UCPD a volte rifiuta a PowerShell le chiavi delle associazioni: reg.exe passa.
        $k = $chiave -replace '^HKLM:', 'HKLM' -replace '^HKCU:', 'HKCU'
        if ($nome -eq '(default)') { $out = & reg.exe add $k /ve /d "$valore" /f 2>&1 }
        else { $out = & reg.exe add $k /v $nome /d "$valore" /f 2>&1 }
        if ($LASTEXITCODE -ne 0) { throw "Registro: impossibile scrivere $chiave ($out)" }
        Scrivi "Nota: $chiave scritta con reg.exe (PowerShell bloccato da Windows)."
    }
}

function Notifica-Shell {
    try {
        if (-not ('ApDocs.Shell' -as [type])) {
            Add-Type -Namespace ApDocs -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, int f, System.IntPtr a, System.IntPtr b);'
        }
        [ApDocs.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    } catch { }
}

function Registra {
    $icoApp = Join-Path $Dir 'brand\apdocs.ico'
    $icoDoc = Join-Path $Dir 'brand\apdocs-document.ico'
    $cmd = "`"$Exe`" `"%1`""
    # tipo di documento apDocs
    Imposta "$Classes\$ProgId" '(default)' 'Documento PDF'
    Imposta "$Classes\$ProgId" 'FriendlyTypeName' 'Documento PDF'
    Imposta "$Classes\$ProgId\DefaultIcon" '(default)' "`"$icoDoc`""
    Imposta "$Classes\$ProgId\shell\open\command" '(default)' $cmd
    Imposta "$Classes\$ProgId\Application" 'ApplicationName' $Nome
    Imposta "$Classes\$ProgId\Application" 'ApplicationIcon' "`"$icoApp`""
    Imposta "$Classes\$ProgId\Application" 'ApplicationCompany' 'apStore.net di Piccolo Antonio'
    Imposta "$Classes\$ProgId\Application" 'ApplicationDescription' 'Lettore ed editor PDF'
    Imposta "$Classes\.pdf\OpenWithProgids" $ProgId ''
    # La scelta fatta dalle Impostazioni di Windows punta all'eseguibile: anche li' nome e icona apDocs.
    Imposta "$Classes\Applications\pdfcraft.exe" 'FriendlyAppName' $Nome
    Imposta "$Classes\Applications\pdfcraft.exe\DefaultIcon" '(default)' "`"$icoDoc`""
    Imposta "$Classes\Applications\pdfcraft.exe\shell\open\command" '(default)' $cmd
    Imposta "$Classes\Applications\pdfcraft.exe\SupportedTypes" '.pdf' ''
    # programma registrato (compare in App predefinite)
    Imposta "$Hive\Software\$Nome\Capabilities" 'ApplicationName' $Nome
    Imposta "$Hive\Software\$Nome\Capabilities" 'ApplicationDescription' 'Lettore ed editor PDF'
    Imposta "$Hive\Software\$Nome\Capabilities" 'ApplicationIcon' "`"$icoApp`""
    Imposta "$Hive\Software\$Nome\Capabilities\FileAssociations" '.pdf' $ProgId
    Imposta "$Hive\Software\RegisteredApplications" $Nome "Software\$Nome\Capabilities"
    # voce in App installate, con disinstallazione
    $u = "$Hive\Software\Microsoft\Windows\CurrentVersion\Uninstall\$Nome"
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $extra = ''
    if ($Utente) { $extra = ' -Utente' }
    Imposta $u 'DisplayName' $Nome
    Imposta $u 'DisplayVersion' (Versione-Installata)
    Imposta $u 'Publisher' 'apStore.net di Piccolo Antonio'
    Imposta $u 'DisplayIcon' $icoApp
    Imposta $u 'InstallLocation' $Dir
    Imposta $u 'URLInfoAbout' 'https://antoniopiccolo.com'
    Imposta $u 'UninstallString' "`"$ps`" -NoProfile -ExecutionPolicy Bypass -File `"$Dir\installa.ps1`" -Disinstalla$extra"
    New-ItemProperty -Path $u -Name 'NoModify' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $u -Name 'NoRepair' -Value 1 -PropertyType DWord -Force | Out-Null
    # collegamento nel menu Start
    $lnk = Join-Path $StartMenu "$Nome.lnk"
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
    $s.TargetPath = $Exe
    $s.WorkingDirectory = $AppDir
    $s.IconLocation = "$icoApp,0"
    $s.Description = "$Nome - lettore ed editor PDF (basato su PdfCraft)"
    $s.Save()
    # Scansiona in apDocs: finestra di scansione di Windows, PDF ricercabile, apertura in apDocs
    if (Test-Path (Join-Path $Dir 'scansiona.ps1')) {
        $s = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $StartMenu "Scansiona in $Nome.lnk"))
        $s.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $s.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Dir\scansiona.ps1`""
        $s.WorkingDirectory = $Dir
        $s.IconLocation = "$icoApp,0"
        $s.Description = "Scansiona dallo scanner e apri il PDF in $Nome"
        $s.Save()
    }
    Notifica-Shell
    Scrivi 'Registrato come programma per i PDF, con collegamento nel menu Start.'
}

function Pianifica-Aggiornamento {
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $extra = ''
    if ($Utente) { $extra = ' -Utente' }
    $arg = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$Dir\installa.ps1`" -Aggiorna -Silenzioso$extra"
    $azione = New-ScheduledTaskAction -Execute $ps -Argument $arg
    $trigger = @((New-ScheduledTaskTrigger -Daily -At '12:30'), (New-ScheduledTaskTrigger -AtLogOn))
    $impost = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    if ($Utente) {
        $princ = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
        $nomeT = "$NomeAttivita ($env:USERNAME)"
    } else {
        $princ = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $nomeT = $NomeAttivita
    }
    Register-ScheduledTask -TaskName $nomeT -Action $azione -Trigger $trigger -Settings $impost -Principal $princ -Description 'Controlla e installa le versioni nuove di apDocs.' -Force | Out-Null
    Scrivi "Aggiornamento automatico pianificato: ogni giorno alle 12:30 e a ogni accesso."
}

function Rimuovi-AltriPdf {
    $tieni = 'infocert|dike|aruba|firma|sign|namirial|bit4id|print to pdf|apdocs'
    $cerca = 'pdf|acrobat|foxit|nitro|sumatra|xchange|pdf24|pdfelement|soda pdf'
    $chiavi = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $trovati = Get-ItemProperty $chiavi -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -match $cerca -and $_.DisplayName -notmatch $tieni -and -not $_.SystemComponent }
    $appx = Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'acrobat|foxit|xodo|pdf' -and $_.Name -notmatch 'apdocs' -and -not $_.IsFramework }
    if (-not $trovati -and -not $appx) { Scrivi 'Nessun altro programma PDF da togliere.'; return }
    Scrivi 'Altri programmi PDF trovati:'
    foreach ($t in $trovati) { Scrivi "  - $($t.DisplayName) $($t.DisplayVersion)" }
    foreach ($a in $appx) { Scrivi "  - $($a.Name) (app dello Store)" }
    if (-not $Silenzioso) {
        $r = Read-Host 'Li disinstallo? (s/N)'
        if ($r -notmatch '^[sS]') { Scrivi 'Nessun programma tolto.'; return }
    }
    foreach ($t in $trovati) {
        try {
            if ($t.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$' -and "$($t.UninstallString)" -match 'msiexec') {
                $p = Start-Process msiexec.exe -ArgumentList '/x', $t.PSChildName, '/qn', '/norestart' -Wait -PassThru
                Scrivi "Tolto $($t.DisplayName) (codice $($p.ExitCode))."
            } elseif ($t.QuietUninstallString) {
                $p = Start-Process cmd.exe -ArgumentList '/c', $t.QuietUninstallString -Wait -PassThru
                Scrivi "Tolto $($t.DisplayName) (codice $($p.ExitCode))."
            } else {
                Scrivi "Da togliere a mano: $($t.DisplayName) (non ha una disinstallazione silenziosa)."
            }
        } catch { Scrivi "Errore togliendo $($t.DisplayName): $($_.Exception.Message)" }
    }
    foreach ($a in $appx) {
        try { $a | Remove-AppxPackage; Scrivi "Tolto $($a.Name)." } catch { Scrivi "Errore togliendo $($a.Name): $($_.Exception.Message)" }
    }
}

function Apri-SceltaPredefinito {
    if ($Silenzioso) { return }
    $q = 'registeredAppMachine'
    if ($Utente) { $q = 'registeredAppUser' }
    Write-Host ''
    Write-Host "Ultimo passo: nelle Impostazioni scegli $Nome per i file .pdf (Windows non permette di farlo in automatico)."
    Start-Process "ms-settings:defaultapps?$q=$Nome"
}

function Installa {
    $man = Leggi-Manifesto
    $pac = Prepara-Pacchetto $man
    Get-Process pdfcraft -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "$AppDir*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path "$Dir\brand" -Force | Out-Null
    foreach ($f in 'apdocs.ico', 'apdocs-document.ico') { Scarica-Verificato "brand/$f" (Join-Path $Dir "brand\$f") }
    Scarica-Verificato 'scansiona.ps1' (Join-Path $Dir 'scansiona.ps1')
    Applica-Pacchetto $pac $man.versione
    Copia-Script
    Registra
    Pianifica-Aggiornamento
    if ($RimuoviAltri) { Rimuovi-AltriPdf }
    Scrivi "$Nome $($man.versione) installato in $Dir."
    Apri-SceltaPredefinito
}

function Aggiorna-Ora {
    $attuale = Versione-Installata
    if (-not $attuale) { throw "$Nome non risulta installato in $Dir." }
    $f = Join-Path $Dir 'sorgente.txt'
    if (-not $SorgenteEsplicita -and (Test-Path $f)) { $script:Sorgente = (Get-Content $f -Raw).Trim() }
    $man = Leggi-Manifesto
    if ([version]$man.versione -le [version]$attuale) { Scrivi "Nessun aggiornamento: installata $attuale, disponibile $($man.versione)."; return }
    if (Get-Process pdfcraft -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "$AppDir*" }) {
        Scrivi "Disponibile $($man.versione), ma $Nome e' aperto: riprovo al prossimo giro."
        return
    }
    # Prima lo script: una versione nuova puo' portare regole nuove, ed e' lo script nuovo
    # (verificato con la firma del manifesto) a doverla installare.
    $mio = Join-Path $Dir 'installa.ps1'
    $prima = if (Test-Path $mio) { (Get-FileHash $mio -Algorithm SHA256).Hash.ToLower() } else { '' }
    try { Scarica-Verificato 'installa.ps1' $mio } catch { Scrivi "Script non aggiornato: $($_.Exception.Message)" }
    $dopo = if (Test-Path $mio) { (Get-FileHash $mio -Algorithm SHA256).Hash.ToLower() } else { '' }
    if ($dopo -and $dopo -ne $prima -and -not $DopoScript) {
        Scrivi 'Script aggiornato: continua lo script nuovo.'
        $argomenti = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $mio, '-Aggiorna', '-DopoScript', '-Sorgente', $Sorgente)
        if ($Silenzioso) { $argomenti += '-Silenzioso' }
        if ($Utente) { $argomenti += '-Utente' }
        & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" @argomenti
        return
    }
    $pac = Prepara-Pacchetto $man
    Applica-Pacchetto $pac $man.versione
    foreach ($f in 'apdocs.ico', 'apdocs-document.ico') { try { Scarica-Verificato "brand/$f" (Join-Path $Dir "brand\$f") } catch { Scrivi "Icona non aggiornata: $($_.Exception.Message)" } }
    try { Scarica-Verificato 'scansiona.ps1' (Join-Path $Dir 'scansiona.ps1') } catch { Scrivi "Scansione non aggiornata: $($_.Exception.Message)" }
    Registra
    Scrivi "Aggiornato da $attuale a $($man.versione)."
}

function Disinstalla-Ora {
    Get-Process pdfcraft -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "$AppDir*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    $nomeT = $NomeAttivita
    if ($Utente) { $nomeT = "$NomeAttivita ($env:USERNAME)" }
    Unregister-ScheduledTask -TaskName $nomeT -Confirm:$false -ErrorAction SilentlyContinue
    Remove-ItemProperty "$Classes\.pdf\OpenWithProgids" -Name $ProgId -ErrorAction SilentlyContinue
    Remove-ItemProperty "$Hive\Software\RegisteredApplications" -Name $Nome -ErrorAction SilentlyContinue
    foreach ($k in "$Classes\$ProgId", "$Classes\Applications\pdfcraft.exe", "$Hive\Software\$Nome", "$Hive\Software\Microsoft\Windows\CurrentVersion\Uninstall\$Nome") {
        if (-not (Test-Path $k)) { continue }
        try { Remove-Item $k -Recurse -Force -ErrorAction Stop }
        catch {
            # UCPD (Windows 11) a volte blocca PowerShell sulle chiavi delle associazioni: reg.exe passa.
            $r = $k -replace '^HKLM:', 'HKLM' -replace '^HKCU:', 'HKCU'
            & reg.exe delete $r /f 2>&1 | Out-Null
            if (Test-Path $k) { Scrivi "Non riesco a togliere $k" }
        }
    }
    Remove-Item (Join-Path $StartMenu "$Nome.lnk"), (Join-Path $StartMenu "Scansiona in $Nome.lnk") -Force -ErrorAction SilentlyContinue
    Notifica-Shell
    # La cartella si toglie dopo che questo script (che sta li' dentro) e' terminato.
    $cmd = "ping -n 3 127.0.0.1 >nul & rmdir /s /q `"$Dir`""
    Start-Process cmd.exe -ArgumentList '/c', $cmd -WindowStyle Hidden
    if (-not $Silenzioso) { Write-Host "$Nome disinstallato. I PDF si apriranno con il programma scelto in Windows." }
}

# --- amministratore: per tutti gli utenti lo script si rilancia elevato ---
if (-not $Utente -and -not (Test-Admin)) {
    if ($Silenzioso) { throw "Per installare per tutti gli utenti servono i permessi di amministratore (oppure usa -Utente)." }
    $file = $PSCommandPath
    if (-not $file) {
        $file = Join-Path $env:TEMP 'apdocs-installa.ps1'
        Leggi-Manifesto | Out-Null
        Scarica-Verificato 'installa.ps1' $file
    }
    $argomenti = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$file`"", '-Sorgente', "`"$Sorgente`"", '-Attendi')
    foreach ($sw in 'Aggiorna', 'Disinstalla', 'RimuoviAltri') { if ((Get-Variable $sw -ValueOnly)) { $argomenti += "-$sw" } }
    Write-Host 'Servono i permessi di amministratore: conferma nella finestra di Windows.'
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argomenti -Wait
    return
}

try {
    if ($Disinstalla) { Disinstalla-Ora }
    elseif ($Aggiorna) { Aggiorna-Ora }
    else { Installa }
} catch {
    Scrivi "ERRORE: $($_.Exception.Message) (riga $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim()))"
    if (-not $Silenzioso) { Write-Host ''; Write-Host 'Qualcosa non e'' andato: scrivi a info@apstore.net con il file' $Registro -ForegroundColor Yellow }
    if ($Silenzioso) { exit 1 }
}
if ($Attendi -and -not $Silenzioso) { Read-Host 'Premi Invio per chiudere' | Out-Null }
