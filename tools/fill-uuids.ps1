# Заменяет каждый маркер @NEWUUID@ в XML-исходниках расширения на новый уникальный GUID.
# Запускается один раз после добавления новых объектов: уже выданные GUID не трогаются,
# поэтому идентификаторы метаданных (а значит, и данные в базе) остаются стабильными.
param(
    [string]$Root = (Join-Path $PSScriptRoot '..\src\cfe')
)

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
Get-ChildItem -Path $Root -Recurse -Include *.xml | ForEach-Object {
    $text = [System.IO.File]::ReadAllText($_.FullName)
    if ($text.Contains('@NEWUUID@')) {
        $result = [regex]::Replace($text, '@NEWUUID@', { param($m) [guid]::NewGuid().ToString() })
        [System.IO.File]::WriteAllText($_.FullName, $result, $utf8Bom)
        Write-Host "UUID: $($_.FullName)"
    }
}
