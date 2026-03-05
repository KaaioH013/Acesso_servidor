$ErrorActionPreference = 'Stop'

$logFile = 'exports/_gerar_excel_diretor.log'
"Inicio: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Set-Content -Path $logFile

$src = Get-ChildItem exports -Filter 'relatorio_consumo_rotor_estator_24m_*.xlsx' |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $src) {
    throw 'Nao encontrei arquivo fonte relatorio_consumo_rotor_estator_24m_*.xlsx em exports/.'
}

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$out = Join-Path (Resolve-Path 'exports') ("relatorio_diretor_2_abas_{0}.xlsx" -f $ts)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wbSrc = $null
$wbOut = $null

function Get-HeaderMap($ws) {
    $map = @{}
    $lastCol = $ws.UsedRange.Columns.Count
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = [string]$ws.Cells.Item(1, $c).Text
        if ($h) { $map[$h] = $c }
    }
    return $map
}

try {
    $wbSrc = $excel.Workbooks.Open($src)
    $wsBase = $wbSrc.Worksheets.Item('Base_Item')
    $wsCons = $wbSrc.Worksheets.Item('Consumo_Mensal')

    $wbOut = $excel.Workbooks.Add()
    while ($wbOut.Worksheets.Count -gt 1) { $wbOut.Worksheets.Item($wbOut.Worksheets.Count).Delete() }

    $wsMin = $wbOut.Worksheets.Item(1)
    $wsMin.Name = 'Minimo_Sugerido_Anual'

    $wsSaz = $wbOut.Worksheets.Add()
    $wsSaz.Name = 'Sazonalidade'

    $wsBase.UsedRange.Copy($wsMin.Range('A1'))

    $keep = @(
        'MATERIAL',
        'DESCRICAO',
        'TIPO_ITEM',
        'ESTOQUE_ATUAL',
        'VMINIMO_BASE',
        'MEDIA_MENSAL_QTDE',
        'MINIMO_SUGERIDO_FINAL',
        'GAP_REPOSICAO',
        'STATUS'
    )

    $lastColMin = $wsMin.UsedRange.Columns.Count
    for ($c = $lastColMin; $c -ge 1; $c--) {
        $h = [string]$wsMin.Cells.Item(1, $c).Text
        if ($h -and ($keep -notcontains $h)) {
            $wsMin.Columns.Item($c).Delete() | Out-Null
        }
    }

    $wsMin.UsedRange.EntireColumn.AutoFit() | Out-Null
    $wsMin.Rows.Item(1).Font.Bold = $true

    $consHdr = Get-HeaderMap $wsCons
    foreach ($need in @('ANO', 'MES', 'CONSUMO_QTDE')) {
        if (-not $consHdr.ContainsKey($need)) {
            throw "Coluna $need nao encontrada em Consumo_Mensal"
        }
    }

    $lastRowCons = $wsCons.UsedRange.Rows.Count
    $totals = @{}

    for ($r = 2; $r -le $lastRowCons; $r++) {
        $anoTxt = [string]$wsCons.Cells.Item($r, $consHdr['ANO']).Text
        $mesTxt = [string]$wsCons.Cells.Item($r, $consHdr['MES']).Text
        $qtdTxt = [string]$wsCons.Cells.Item($r, $consHdr['CONSUMO_QTDE']).Text

        if ([string]::IsNullOrWhiteSpace($anoTxt) -or [string]::IsNullOrWhiteSpace($mesTxt)) { continue }

        $ano = [int]$anoTxt
        $mes = [int]$mesTxt
        $qtd = 0.0
        if (-not [double]::TryParse(($qtdTxt -replace ',', '.'), [ref]$qtd)) { continue }

        $k = "{0}|{1}" -f $ano, $mes
        if (-not $totals.ContainsKey($k)) { $totals[$k] = 0.0 }
        $totals[$k] += $qtd
    }

    $mesNomes = @{1='Jan';2='Fev';3='Mar';4='Abr';5='Mai';6='Jun';7='Jul';8='Ago';9='Set';10='Out';11='Nov';12='Dez'}

    $wsSaz.Cells.Item(1, 1).Value2 = 'MES'
    $wsSaz.Cells.Item(1, 2).Value2 = 'MES_NOME'
    $wsSaz.Cells.Item(1, 3).Value2 = 'CONSUMO_MEDIO_QTDE'

    for ($mes = 1; $mes -le 12; $mes++) {
        $vals = @()
        foreach ($k in $totals.Keys) {
            $parts = $k.Split('|')
            $m = [int]$parts[1]
            if ($m -eq $mes) { $vals += [double]$totals[$k] }
        }

        $media = if ($vals.Count -gt 0) { ($vals | Measure-Object -Average).Average } else { 0 }
        $row = $mes + 1

        $wsSaz.Cells.Item($row, 1).Value2 = [string]$mes
        $wsSaz.Cells.Item($row, 2).Value2 = [string]$mesNomes[$mes]
        $wsSaz.Cells.Item($row, 3).Value2 = [string]([math]::Round($media, 0))
    }

    $wsSaz.UsedRange.EntireColumn.AutoFit() | Out-Null
    $wsSaz.Rows.Item(1).Font.Bold = $true

    $wbOut.SaveAs($out, 51)
    "Saida: $out" | Set-Content -Path $logFile
}
catch {
    "ERRO: $($_.Exception.Message)" | Set-Content -Path $logFile
    throw
}
finally {
    if ($wbOut -ne $null) { $wbOut.Close($true) }
    if ($wbSrc -ne $null) { $wbSrc.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }

    if ($wbOut -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbOut) | Out-Null }
    if ($wbSrc -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbSrc) | Out-Null }
    if ($excel -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
