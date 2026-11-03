# Windows built-in OCR (Korean) -> words with bounding boxes, one JSON line per image
# powershell -File ocr_words.ps1 -List <list.txt (UTF-8, one image path per line)> -Out <result.jsonl> [-Lang ko]
param([string]$List, [string]$Out, [string]$Lang = 'ko')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait() | Out-Null; $task.Result }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language $Lang))
$sw = New-Object System.IO.StreamWriter($Out, $false, (New-Object System.Text.UTF8Encoding($false)))
if (-not $engine) { $sw.WriteLine('{"error":"no-ocr-language"}'); $sw.Close(); exit }
foreach ($path in [System.IO.File]::ReadAllLines($List, [System.Text.Encoding]::UTF8)) {
  if (-not $path.Trim()) { continue }
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $words = New-Object System.Collections.ArrayList
    foreach ($line in $result.Lines) {
      foreach ($w in $line.Words) {
        $b = $w.BoundingRect
        [void]$words.Add(@([math]::Round($b.X, 1), [math]::Round($b.Y, 1), [math]::Round($b.Width, 1), [math]::Round($b.Height, 1), $w.Text))
      }
    }
    $stream.Dispose()
    $obj = @{ file = $path; w = $bitmap.PixelWidth; h = $bitmap.PixelHeight; words = $words }
  } catch {
    $obj = @{ file = $path; error = $_.Exception.Message; words = @() }
  }
  $sw.WriteLine((ConvertTo-Json $obj -Compress -Depth 5))
  $sw.Flush()
}
$sw.Close()
