# instala el módulo ImportExcel (sólo la primera vez)
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
Install-Module -Name ImportExcel -Scope CurrentUser -Force

$inputPath  = "C:\Temp\Mexico\Input-Reportes"
$outputPath = $inputPath

Get-ChildItem -Path $inputPath -Filter "*.xls*" | ForEach-Object {
    $excelFile         = $_.FullName
    $fileNameNoExt     = [IO.Path]::GetFileNameWithoutExtension($_.Name)
    $outputCsv         = Join-Path $outputPath "$fileNameNoExt.csv"

    try {
        # 1) Leemos TODO sin cabecera
        $data = Import-Excel -Path $excelFile -NoHeader

        # 2) Convertimos a texto CSV pipe-delimited, sin línea de tipo,
        #    y quitamos todas las comillas (para que no "escapen" nada).
        $csvLines = $data |
            ConvertTo-Csv -NoTypeInformation -Delimiter '|' |
            Select-Object -Skip 1 |                     # quitamos la cabecera automática
            ForEach-Object { $_ -Replace '"', '' }      # eliminamos comillas sobrantes

        # 3) Guardamos en UTF-8
        $csvLines | Out-File -FilePath $outputCsv -Encoding UTF8

        Write-Host "✅ Convertido: $excelFile → $outputCsv"
    }
    catch {
        Write-Warning "⚠ Error al procesar: $excelFile → $_"
    }
}
