<#
apDocs - Scansiona: acquisisce dallo scanner con la finestra di Windows (WIA), crea un PDF con apDocs,
lo rende ricercabile (OCR), lo salva in Documenti\Scansioni e lo apre in apDocs.

  Dal menu Start: "Scansiona in apDocs" (lo crea installa.ps1).
  .\scansiona.ps1 [-SenzaOcr] [-Cartella <percorso>] [-CambiaScanner]

Dal caricatore automatico prende tutti i fogli; dal piano chiede se c'e' un'altra pagina.
Lo scanner scelto la prima volta viene ricordato (%APPDATA%\apDocs\scanner.txt).
#>
param(
    [switch]$SenzaOcr,
    [string]$Cartella = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Scansioni'),
    [switch]$CambiaScanner
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

$JPEG = '{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}'
$WIA_PAPER_EMPTY = -2145320957        # 0x80210003: il caricatore e' vuoto
$WIA_DOCUMENT_HANDLING_SELECT = 3088  # 1 = caricatore, 2 = piano
$WIA_HORIZONTAL_RESOLUTION = 6147

$cli = Join-Path $PSScriptRoot 'app\pdfcraft-cli.exe'
$app = Join-Path $PSScriptRoot 'app\pdfcraft.exe'
$memoria = Join-Path $env:APPDATA 'apDocs\scanner.txt'

function Avviso([string]$testo, [string]$icona = 'Information') {
    [void][Windows.Forms.MessageBox]::Show($testo, 'Scansiona in apDocs', 'OK', $icona)
}
function Domanda([string]$testo) {
    return [Windows.Forms.MessageBox]::Show($testo, 'Scansiona in apDocs', 'YesNo', 'Question') -eq 'Yes'
}
function Proprieta($oggetto, [int]$id) {
    foreach ($p in $oggetto.Properties) { if ($p.PropertyID -eq $id) { return $p.Value } }
    return $null
}

$lavoro = Join-Path $env:TEMP ('apdocs-scansione-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $lavoro | Out-Null
try {
    if (-not (Test-Path $cli)) { throw "apDocs non trovato in $PSScriptRoot." }
    $dlg = New-Object -ComObject WIA.CommonDialog

    # scanner: quello ricordato, altrimenti lo si sceglie (e lo si ricorda)
    $dev = $null
    if (-not $CambiaScanner -and (Test-Path $memoria)) {
        $id = (Get-Content $memoria -Raw).Trim()
        $dm = New-Object -ComObject WIA.DeviceManager
        foreach ($info in $dm.DeviceInfos) { if ($info.DeviceID -eq $id) { try { $dev = $info.Connect() } catch { $dev = $null } } }
    }
    if (-not $dev) {
        $dev = $dlg.ShowSelectDevice(1, $true, $false)   # 1 = scanner; sceglie da solo se ce n'e' uno
        if (-not $dev) { return }
        New-Item -ItemType Directory -Path (Split-Path $memoria) -Force | Out-Null
        Set-Content -Path $memoria -Value $dev.DeviceID -Encoding ASCII
    }

    # finestra di Windows: origine (piano o caricatore), colore, risoluzione, anteprima
    $items = $dlg.ShowSelectItems($dev, 1, 0, $true, $true, $false)
    if (-not $items -or $items.Count -lt 1) { return }
    $item = $items.Item(1)
    $dalCaricatore = ((Proprieta $dev $WIA_DOCUMENT_HANDLING_SELECT) -band 1) -eq 1
    $dpi = Proprieta $item $WIA_HORIZONTAL_RESOLUTION

    $immagini = New-Object System.Collections.Generic.List[string]
    while ($true) {
        try {
            $img = $dlg.ShowTransfer($item, $JPEG, $false)
        } catch [Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -eq $WIA_PAPER_EMPTY -and $immagini.Count -gt 0) { break }
            if ($_.Exception.HResult -eq $WIA_PAPER_EMPTY) { throw 'Il caricatore e'' vuoto: metti i fogli e riprova.' }
            throw
        }
        if (-not $img) { break }
        $f = Join-Path $lavoro ('pagina-{0:D3}.jpg' -f ($immagini.Count + 1))
        $img.SaveFile($f)
        $immagini.Add($f)
        if ($dalCaricatore) { continue }
        if (-not (Domanda "Pagina $($immagini.Count) acquisita. Scansionare un'altra pagina dal piano?")) { break }
    }
    if ($immagini.Count -eq 0) { return }

    # PDF con apDocs (riga di comando, confinata alla cartella di lavoro)
    $nome = 'Scansione {0:yyyy-MM-dd HHmmss}.pdf' -f (Get-Date)
    $pdf = Join-Path $lavoro 'scansione.pdf'
    $crea = [ordered]@{ from = 'images'; paths = @($immagini) }
    if ($dpi) { $crea.dpi = [double]$dpi }
    $passi = @(@{ tool = 'doc_create'; args = $crea })
    if (-not $SenzaOcr) { $passi += @{ tool = 'ocr_recognize'; args = @{ doc = 1 } } }
    $passi += @{ tool = 'doc_save'; args = @{ doc = 1; path = $pdf } }
    $script = Join-Path $lavoro 'passi.json'
    [IO.File]::WriteAllText($script, (ConvertTo-Json -InputObject $passi -Depth 5), (New-Object Text.UTF8Encoding $false))
    $out = & $cli run --script $script --root $lavoro 2>&1
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $pdf)) { throw "apDocs non ha creato il PDF: $($out | Out-String)" }

    New-Item -ItemType Directory -Path $Cartella -Force | Out-Null
    $dest = Join-Path $Cartella $nome
    Move-Item $pdf $dest
    Start-Process -FilePath $app -ArgumentList "`"$dest`""
} catch {
    Avviso "La scansione non e' riuscita.`n`n$($_.Exception.Message)`n`nSe il problema resta, scrivi a supporto@apstore.net." 'Warning'
} finally {
    Remove-Item $lavoro -Recurse -Force -ErrorAction SilentlyContinue
}
