function Inject-Docx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Jalur lengkap menuju file .docx asli")]
        [string]$Path,

        [Parameter(Mandatory = $true, HelpMessage = "Blok kode HTML/JS yang akan disuntikkan")]
        [string]$HtmlBlock,

        [Parameter(Mandatory = $false, HelpMessage = "Nama atau jalur file hasil akhir")]
        [string]$DestinationName = "output.docx"
    )

    # 1. Validasi keberadaan file sumber
    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Error "File sumber tidak ditemukan pada jalur: $Path"
        return
    }

    # Resolve path absolut untuk menghindari masalah direktori relatif
    $resolvedPath = (Resolve-Path -Path $Path).Path
    $workingDir = Split-Path -Path $resolvedPath -Parent

    # Tentukan lokasi output akhir (mendukung nama file saja atau full path)
    if ([System.IO.Path]::IsPathRooted($DestinationName)) {
        $finalOutputPath = $DestinationName
    } else {
        $finalOutputPath = Join-Path -Path $workingDir -ChildPath $DestinationName
    }

    # Buat direktori dan nama file sementara yang unik agar tidak bentrok dengan file yang sudah ada
    $uniqueId = [guid]::NewGuid().ToString("N")
    $tempZipPath = Join-Path -Path $env:TEMP -ChildPath "docx_temp_$uniqueId.zip"
    $tempExtractDir = Join-Path -Path $env:TEMP -ChildPath "docx_extract_$uniqueId"

    try {
        # 2. Salin dokumen asli ke folder TEMP sebagai arsip ZIP (file asli tetap aman)
        Copy-Item -Path $resolvedPath -Destination $tempZipPath -Force
        Expand-Archive -Path $tempZipPath -DestinationPath $tempExtractDir -Force

        # 3. Periksa dan muat file document.xml
        $documentXmlPath = Join-Path -Path $tempExtractDir -ChildPath "word\document.xml"
        if (-not (Test-Path -Path $documentXmlPath)) {
            Write-Error "Struktur dokumen tidak valid. File word\document.xml tidak ditemukan."
            return
        }

        [xml]$xml = Get-Content -Path $documentXmlPath -Encoding UTF8
        $injectionCount = 0

        # 4. Pindai seluruh node XML untuk mencari atribut embeddedHtml
        $xml.SelectNodes("//*") | ForEach-Object {
            if ($_.HasAttribute("embeddedHtml")) {
                $_.SetAttribute("embeddedHtml", $HtmlBlock)
                $injectionCount++
            }
        }

        # Beri peringatan jika dokumen tidak memiliki video daring sama sekali
        if ($injectionCount -eq 0) {
            Write-Warning "Tidak ditemukan atribut 'embeddedHtml' di dalam dokumen. Pastikan dokumen sudah memiliki video daring (Insert > Online Video)."
            return
        }

        # Simpan perubahan XML dengan encoding UTF-8
        $xml.Save($documentXmlPath)

        # 5. Kemas ulang folder menjadi arsip ZIP baru dan simpan ke lokasi tujuan
        $newZipPath = Join-Path -Path $env:TEMP -ChildPath "docx_repacked_$uniqueId.zip"
        Compress-Archive -Path "$tempExtractDir\*" -DestinationPath $newZipPath -Force

        # Pindahkan dan timpa file tujuan jika sudah ada
        Move-Item -Path $newZipPath -Destination $finalOutputPath -Force

        Write-Host "[+] Berhasil! Injeksi diterapkan pada $injectionCount tag embeddedHtml." -ForegroundColor Green
        Write-Host "[+] Dokumen hasil disimpan di: $finalOutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Terjadi kesalahan saat memproses dokumen: $_"
    }
    finally {
        # 6. Bersihkan seluruh file dan folder sementara di direktori TEMP
        if (Test-Path -Path $tempZipPath) { Remove-Item -Path $tempZipPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -Path $tempExtractDir) { Remove-Item -Path $tempExtractDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
