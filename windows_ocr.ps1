param(
    [Parameter(Mandatory=$true)]
    [string]$ImagePath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Await($AsyncOperation, [Type]$ResultType) {
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq "AsTask" -and
            $_.IsGenericMethod -and
            $_.GetParameters().Count -eq 1
        } |
        Select-Object -First 1

    $generic = $method.MakeGenericMethod($ResultType)
    $task = $generic.Invoke($null, @($AsyncOperation))
    $task.Wait()
    return $task.Result
}

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]

    $fullPath = [System.IO.Path]::GetFullPath($ImagePath)
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        throw "Windows OCR 엔진을 생성하지 못했습니다. Windows 언어 팩의 OCR 기능을 확인하세요."
    }

    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $text = $result.Text

    # 시간, 날짜, 페이지 표기 등을 제거한 뒤 1~45 숫자만 추출
    $clean = [regex]::Replace($text, '\b\d{1,2}:\d{2}\b', ' ')
    $clean = [regex]::Replace($clean, '\b\d{4}[./-]\d{1,2}[./-]\d{1,2}\b', ' ')
    $clean = [regex]::Replace($clean, '\b\d+\s*/\s*\d+\b', ' ')

    $numbers = @()
    foreach ($m in [regex]::Matches($clean, '(?<!\d)\d{1,2}(?!\d)')) {
        $n = [int]$m.Value
        if ($n -ge 1 -and $n -le 45) {
            $numbers += $n
        }
    }

    @{
        ok = $true
        text = $text
        numbers = $numbers
    } | ConvertTo-Json -Compress -Depth 4
    exit 0
}
catch {
    @{
        ok = $false
        error = $_.Exception.Message
        numbers = @()
    } | ConvertTo-Json -Compress -Depth 4
    exit 1
}
