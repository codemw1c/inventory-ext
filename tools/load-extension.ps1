# Загружает расширение из XML-исходников в информационную базу и обновляет ее.
# Конфигуратор с этой базой должен быть закрыт.
param(
    [string]$InfoBase  = 'D:\YandexDisk\1C Programming\DataBases\db001',
    [string]$Platform  = 'C:\Program Files (x86)\1cv8\8.3.27.2342\bin\1cv8.exe',
    [string]$Extension = 'СкладскойУчет',
    [string]$Source    = (Join-Path $PSScriptRoot '..\src\cfe')
)

$Source = (Resolve-Path $Source).Path
$log = Join-Path $env:TEMP 'cfe-load.log'

function Invoke-Designer([string[]]$Arguments) {
    Remove-Item $log -ErrorAction SilentlyContinue
    $all = @('DESIGNER', '/F', "`"$InfoBase`"", '/DisableStartupDialogs', '/DisableStartupMessages') +
        $Arguments + @('/Out', "`"$log`"")
    $p = Start-Process $Platform -ArgumentList $all -Wait -PassThru
    if (Test-Path $log) { Get-Content $log }
    if ($p.ExitCode -ne 0) { throw "Конфигуратор завершился с кодом $($p.ExitCode)" }
}

Invoke-Designer @('/LoadConfigFromFiles', "`"$Source`"", '-Extension', $Extension)
Invoke-Designer @('/UpdateDBCfg', '-Extension', $Extension)
Write-Host "Расширение $Extension загружено в $InfoBase"
