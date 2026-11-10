$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "[1/4] Consultando API do portal Locarga FTL (CSN)..." -ForegroundColor Cyan
$apiJsonPath = Join-Path $PSScriptRoot "pdfs_ftl\vagoes_api.json"
$pdfsDir = Join-Path $PSScriptRoot "pdfs_ftl"
if (-not (Test-Path $pdfsDir)) { New-Item -ItemType Directory -Path $pdfsDir | Out-Null }

curl.exe --ssl-no-revoke -s -u "SANTOSBR:SANTOS53" "https://locarga-ftl.csn.com.br/locarga/api/vagao/todos?pag=0&pgsize=2000&ord=&dir=" -o $apiJsonPath
$json = Get-Content -Raw $apiJsonPath | ConvertFrom-Json
Write-Host "      -> $($json.content.Count) vagoes retornados pela API." -ForegroundColor Green

Write-Host "[2/4] Cruzando vagoes da planilha 'VTs FTL NEW.xlsx' e verificando PDFs..." -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path "VTs FTL NEW.xlsx"))
$ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
$ss = @()
if ($ssEntry) {
    $sr = New-Object System.IO.StreamReader($ssEntry.Open())
    [xml]$xml = $sr.ReadToEnd()
    $sr.Close()
    foreach ($si in $xml.sst.si) {
        if ($si.t -is [string]) { $ss += $si.t }
        elseif ($si.t.'#text') { $ss += $si.t.'#text' }
        else { $ss += '' }
    }
}
$sEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sr = New-Object System.IO.StreamReader($sEntry.Open())
[xml]$sXml = $sr.ReadToEnd()
$sr.Close()
$zip.Dispose()

$apiByCode = @{}
foreach ($item in $json.content) {
    $codeClean = ($item.codigo_vagao -replace '[^0-9]', '')
    $apiByCode[$codeClean] = $item
    $apiByCode[$codeClean.TrimStart('0')] = $item
}

$manifest = @()
$curlLines = @()
foreach ($r in ($sXml.worksheet.sheetData.row | Select-Object -Skip 1)) {
    $aCell = $r.c | Where-Object { $_.r -match '^A' }
    $bCell = $r.c | Where-Object { $_.r -match '^B' }
    $eCell = $r.c | Where-Object { $_.r -match '^E' }
    $fCell = $r.c | Where-Object { $_.r -match '^F' }
    $gCell = $r.c | Where-Object { $_.r -match '^G' }

    $ser = $aCell.v; if ($aCell.t -eq 's' -and $null -ne $ser) { $ser = $ss[[int]$ser] }
    $veic = $bCell.v; if ($bCell.t -eq 's' -and $null -ne $veic) { $veic = $ss[[int]$veic] }
    $capVol = $eCell.v; if ($eCell.t -eq 's' -and $null -ne $capVol) { $capVol = $ss[[int]$capVol] }
    $capMax = $fCell.v; if ($fCell.t -eq 's' -and $null -ne $capMax) { $capMax = $ss[[int]$capMax] }
    $comp = $gCell.v; if ($gCell.t -eq 's' -and $null -ne $comp) { $comp = $ss[[int]$comp] }

    if ($veic) {
        $c = $veic.ToString().Trim()
        $item = $apiByCode[$c]
        if (-not $item) { $item = $apiByCode[$c.TrimStart('0')] }
        if ($item -and $item.caminho_arquivo) {
            $relPath = $item.caminho_arquivo.Substring(4)
            $url = "https://locarga-ftl.csn.com.br/" + ($relPath -replace ' ', '%20')
            $localFile = Join-Path $pdfsDir "$($item.codigo_vagao)_$($item.nome_arquivo)"
            $manifest += [PSCustomObject]@{
                Row = $r.r
                Ser = $ser
                Veic = $c
                CodigoVagao = $item.codigo_vagao
                SerieVagao = $item.serie_vagao
                DataInsercao = $item.data_insercao
                NomeArquivo = $item.nome_arquivo
                Url = $url
                LocalFile = $localFile
                CapVolPlanilha = $capVol
                CapMaxPlanilha = $capMax
                CompPlanilha = $comp
            }
            if (-not (Test-Path $localFile)) {
                $localEscaped = $localFile -replace '\\', '/'
                $curlLines += "url = `"$url`""
                $curlLines += "output = `"$localEscaped`""
            }
        }
    }
}

$manifestPath = Join-Path $pdfsDir "manifest_ftl.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $manifestPath

if ($curlLines.Count -gt 0) {
    $curlCfg = Join-Path $pdfsDir "curl_urls.txt"
    $curlLines | Set-Content -Encoding ASCII $curlCfg
    Write-Host "      -> Baixando $($curlLines.Count / 2) novos certificados em PDF..." -ForegroundColor Yellow
    curl.exe --ssl-no-revoke --parallel --parallel-max 15 -s -L -u "SANTOSBR:SANTOS53" -K $curlCfg
} else {
    Write-Host "      -> Todos os $($manifest.Count) PDFs ja estao baixados na pasta 'pdfs_ftl'." -ForegroundColor Green
}

Write-Host "[3/4] Processando OCR nos certificados em PDF..." -ForegroundColor Cyan
$ocrCachePath = Join-Path $pdfsDir "ocr_words_368.json"
$existingOcr = @{}
if (Test-Path $ocrCachePath) {
    $loaded = Get-Content -Raw $ocrCachePath | ConvertFrom-Json
    foreach ($it in $loaded) { $existingOcr[$it.File] = $it }
}

$toOcr = @()
foreach ($m in $manifest) {
    $fn = [System.IO.Path]::GetFileName($m.LocalFile)
    if (-not $existingOcr.ContainsKey($fn)) {
        $toOcr += $m.LocalFile
    }
}

if ($toOcr.Count -gt 0) {
    Write-Host "      -> Executando OCR em $($toOcr.Count) novos PDFs..." -ForegroundColor Yellow
    $workerScript = {
        param($filePaths)
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
        function Await($WinRtTask, $ResultType) {
            $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
            $netTask = $asTask.Invoke($null, @($WinRtTask))
            $netTask.Wait(-1) | Out-Null
            return $netTask.Result
        }
        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        $results = @()
        foreach ($fp in $filePaths) {
            $fs = [System.IO.File]::OpenRead($fp)
            $rtStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($fs)
            $pdf = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($rtStream)) ([Windows.Data.Pdf.PdfDocument])
            $page = $pdf.GetPage(0)
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $opts.DestinationWidth = 2200
            $t = $asTaskAction.Invoke($null, @($page.RenderToStreamAsync($stream, $opts)))
            $t.Wait(-1) | Out-Null
            $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
            $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
            $imgW = [double]$bitmap.PixelWidth
            $imgH = [double]$bitmap.PixelHeight
            $ocrResult = Await ($ocrEngine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $fs.Close()
            $stream.Dispose()
            $words = @()
            foreach ($line in $ocrResult.Lines) {
                foreach ($word in $line.Words) {
                    $words += [PSCustomObject]@{
                        T = $word.Text
                        X = [Math]::Round(([double]$word.BoundingRect.X / $imgW), 4)
                        Y = [Math]::Round(([double]$word.BoundingRect.Y / $imgH), 4)
                    }
                }
            }
            $results += [PSCustomObject]@{ File = [System.IO.Path]::GetFileName($fp); W = $imgW; H = $imgH; Words = $words }
        }
        return $results
    }
    $numThreads = 6
    $pool = [runspacefactory]::CreateRunspacePool(1, $numThreads)
    $pool.Open()
    $jobs = @()
    $chunkSize = [Math]::Ceiling($toOcr.Count / $numThreads)
    for ($i = 0; $i -lt $numThreads; $i++) {
        $chunk = $toOcr | Select-Object -Skip ($i * $chunkSize) -First $chunkSize
        if ($chunk.Count -gt 0) {
            $ps = [powershell]::Create().AddScript($workerScript).AddArgument($chunk)
            $ps.RunspacePool = $pool
            $jobs += [PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke() }
        }
    }
    foreach ($j in $jobs) {
        $res = $j.PS.EndInvoke($j.Handle)
        foreach ($r in $res) { $existingOcr[$r.File] = $r }
        $j.PS.Dispose()
    }
    $pool.Close()
    $pool.Dispose()
    $existingOcr.Values | ConvertTo-Json -Depth 5 -Compress | Set-Content -Encoding UTF8 $ocrCachePath
} else {
    Write-Host "      -> Cache de OCR completo para todos os $($manifest.Count) certificados." -ForegroundColor Green
}

Write-Host "[4/4] Atualizando planilha 'VTs FTL NEW.xlsx'..." -ForegroundColor Cyan
powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "pdfs_ftl\build_excel.ps1")
Write-Host "Concluido com sucesso!" -ForegroundColor Green
