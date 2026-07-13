param([string]$Root,[switch]$NoOpen)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new()
if([string]::IsNullOrWhiteSpace($Root)){$Root=Split-Path -Parent $PSScriptRoot}
$Root=[IO.Path]::GetFullPath($Root)
Set-Location -LiteralPath $Root
Write-Host ''
Write-Host '=== TOOL ĐỐI CHIẾU CCDC ===' -ForegroundColor Cyan
$required=@('Bảng kê chứng từ.xlsx','Bảng tổng hợp chi phí chờ phân bổ.xlsx','Check_CCDC 1.xlsx')
foreach($file in $required){if(-not(Test-Path -LiteralPath (Join-Path $Root $file))){Write-Host "THIẾU FILE: $file" -ForegroundColor Red;exit 1}}
$node=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
if(-not(Test-Path -LiteralPath $node)){$cmd=Get-Command node.exe -ErrorAction SilentlyContinue;if($cmd){$node=$cmd.Source}}
if(-not(Test-Path -LiteralPath $node)){Write-Host 'KHÔNG TÌM THẤY BỘ CHẠY NODE. Hãy mở Codex một lần rồi chạy lại tool.' -ForegroundColor Red;exit 1}
$python=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if(-not(Test-Path -LiteralPath $python)){$cmd=Get-Command python.exe -ErrorAction SilentlyContinue;if($cmd){$python=$cmd.Source}}
if(-not(Test-Path -LiteralPath $python)){Write-Host 'KHÔNG TÌM THẤY BỘ CHẠY PYTHON. Hãy mở Codex một lần rồi chạy lại tool.' -ForegroundColor Red;exit 1}
Write-Host 'Đang đọc file nguồn...'
& $node (Join-Path $Root 'TOOL_Check_CCDC\build_stage.mjs')
if($LASTEXITCODE -ne 0){Write-Host 'LỖI KHI TẠO FILE TẠM' -ForegroundColor Red;exit $LASTEXITCODE}
Write-Host 'Đang đối chiếu theo Mã 2421 và 2422...'
& $python (Join-Path $Root 'TOOL_Check_CCDC\finalize_report_fast.py')
if($LASTEXITCODE -ne 0){Write-Host 'LỖI KHI ĐỐI CHIẾU' -ForegroundColor Red;exit $LASTEXITCODE}
$lastFile=Join-Path $Root 'TOOL_Check_CCDC\_last_output.txt'
$output=(Get-Content -LiteralPath $lastFile -Raw -Encoding UTF8).Trim()
$outputPath=Join-Path $Root $output
Remove-Item -LiteralPath (Join-Path $Root 'TOOL_Check_CCDC\_temp_stage.xlsx') -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $Root 'TOOL_Check_CCDC\_temp_stage.xlsx.inspect.ndjson') -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ($outputPath+'.inspect.ndjson') -Force -ErrorAction SilentlyContinue
Write-Host ''
Write-Host "HOÀN TẤT: $output" -ForegroundColor Green
if(-not $NoOpen){Start-Process -FilePath $outputPath}
