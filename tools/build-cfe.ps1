# Собирает файл расширения .cfe из XML-исходников src/cfe.
# Рабочая база не нужна: сборка идет во временной пустой файловой базе.
# С ключом -Verify готовый .cfe дополнительно устанавливается в еще одну чистую базу
# с обновлением структуры — так проверяется, что расширение ставится на «чужую» конфигурацию.
param(
    [string]$Platform  = 'C:\Program Files (x86)\1cv8\8.3.27.2342\bin\1cv8.exe',
    [string]$Extension = 'СкладскойУчет',
    [string]$Source    = (Join-Path $PSScriptRoot '..\src\cfe'),
    [string]$OutDir    = (Join-Path $PSScriptRoot '..\build'),
    [switch]$Verify
)
$ErrorActionPreference = 'Stop'

$Source = (Resolve-Path $Source).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

[xml]$cfg = Get-Content (Join-Path $Source 'Configuration.xml') -Encoding UTF8
$version = $cfg.MetaDataObject.Configuration.Properties.Version
$cfe = Join-Path $OutDir "inventory-ext-$version.cfe"

$work = Join-Path $env:TEMP ('cfe-build-' + [guid]::NewGuid().ToString('N'))
$log = Join-Path $work 'out.log'
New-Item -ItemType Directory -Force $work | Out-Null

function Invoke-1C([string[]]$Arguments) {
    Remove-Item $log -ErrorAction SilentlyContinue
    $p = Start-Process $Platform -ArgumentList ($Arguments + @('/DisableStartupDialogs', '/Out', "`"$log`"")) -Wait -PassThru
    if (Test-Path $log) { Get-Content $log | Where-Object { $_.Trim() } | ForEach-Object { "    $_" } }
    if ($p.ExitCode -ne 0) { throw "1С завершилась с кодом $($p.ExitCode)" }
}

function New-EmptyInfoBase([string]$Path) {
    Invoke-1C @('CREATEINFOBASE', "File=`"$Path`"")
}

try {
    Write-Host "Сборка $cfe"
    $buildBase = Join-Path $work 'build-ib'
    New-EmptyInfoBase $buildBase
    Invoke-1C @('DESIGNER', '/F', "`"$buildBase`"", '/LoadConfigFromFiles', "`"$Source`"", '-Extension', $Extension)
    Invoke-1C @('DESIGNER', '/F', "`"$buildBase`"", '/DumpCfg', "`"$cfe`"", '-Extension', $Extension)

    if ($Verify) {
        Write-Host 'Проверка установки в чистую базу'
        $checkBase = Join-Path $work 'check-ib'
        New-EmptyInfoBase $checkBase
        Invoke-1C @('DESIGNER', '/F', "`"$checkBase`"", '/LoadCfg', "`"$cfe`"", '-Extension', $Extension)
        Invoke-1C @('DESIGNER', '/F', "`"$checkBase`"", '/UpdateDBCfg', '-Extension', $Extension)
        Invoke-1C @('DESIGNER', '/F', "`"$checkBase`"", '/CheckModules', '-Server', '-ThinClient', '-Extension', $Extension)
    }

    Write-Host "Готово: $cfe ($((Get-Item $cfe).Length) байт)"
}
finally {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}
