param(
  [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host "Compilando com PyInstaller (spec)" -ForegroundColor Cyan

if ($Clean) {
  if (Test-Path build) { Remove-Item -Recurse -Force build }
  if (Test-Path dist)  { Remove-Item -Recurse -Force dist }
}

function Ensure-PyInstaller {
  try {
    $null = pyinstaller --version 2>$null
    return
  } catch {
    Write-Host "PyInstaller não encontrado no PATH. Tentando via 'py -m PyInstaller'." -ForegroundColor Yellow
  }
}

Ensure-PyInstaller

# Usa o spec para garantir ícone e modo janela
& py -m PyInstaller --noconfirm --clean 'main.spec'

Write-Host "---"; Write-Host "Concluído. Verifique o executável em 'dist\\main.exe'." -ForegroundColor Green

