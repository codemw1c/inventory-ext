# Дымовой тест расширения через COM-соединение (запускать 32-битным PowerShell под платформу x86):
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -File tools\smoke-test.ps1
# Сценарий: приход 10 -> резерв 4 -> резерв 7 (должен быть отказ) -> расход 4 по резерву ->
# отчет по резервам -> этикетки в PDF.
param(
    [string]$InfoBase = 'D:\YandexDisk\1C Programming\DataBases\db001',
    [string]$OutDir   = (Join-Path $env:TEMP 'cfe-smoke')
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null

$connector = New-Object -ComObject 'V83.COMConnector'
$c = $connector.Connect("File=`"$InfoBase`";")
$failed = 0

function Check($ok, $text) {
    if ($ok) { Write-Host "OK    $text" } else { Write-Host "FAIL  $text"; $script:failed++ }
}
function Messages() {
    $list = $c.ПолучитьСообщенияПользователю($true)
    $res = @(); for ($i = 0; $i -lt $list.Количество(); $i++) { $res += $list.Получить($i).Текст }; $res
}
function Post($doc) {
    try { $doc.Записать($c.РежимЗаписиДокумента.Проведение); $true }
    catch { $false }
}

$stamp = Get-Date -Format 'HHmmss'

$warehouse = $c.Справочники.ск_Склады.СоздатьЭлемент()
$warehouse.Наименование = "Тестовый склад $stamp"
$warehouse.Записать()

$item = $c.Справочники.ск_Номенклатура.СоздатьЭлемент()
$item.Наименование = "Тестовый товар $stamp"
$item.Артикул = "ART-$stamp"
$item.Записать()
$barcode = $item.Штрихкод
Check ($barcode.Length -eq 13 -and $c.ск_ШтрихкодыСервер.ШтрихкодEAN13Корректен($barcode)) "автоматический штрихкод EAN-13: $barcode"
Check ($item.ЕдиницаИзмерения -eq 'шт') "единица измерения по умолчанию"

$income = $c.Документы.ск_ПриходТоваров.СоздатьДокумент()
$income.Дата = Get-Date
$income.Склад = $warehouse.Ссылка
$row = $income.Товары.Добавить(); $row.Номенклатура = $item.Ссылка; $row.Количество = 10; $row.Цена = 150
Check (Post $income) "приход 10 шт. проведен"
Check ($income.СуммаДокумента -eq 1500) "сумма прихода рассчитана: $($income.СуммаДокумента)"

$reserve = $c.Документы.ск_РезервТоваров.СоздатьДокумент()
$reserve.Дата = Get-Date
$reserve.Склад = $warehouse.Ссылка
$row = $reserve.Товары.Добавить(); $row.Номенклатура = $item.Ссылка; $row.Количество = 4
Check (Post $reserve) "резерв 4 шт. проведен"

$over = $c.Документы.ск_РезервТоваров.СоздатьДокумент()
$over.Дата = Get-Date
$over.Склад = $warehouse.Ссылка
$row = $over.Товары.Добавить(); $row.Номенклатура = $item.Ссылка; $row.Количество = 7
$overPosted = Post $over
$msg = (Messages) -join ' | '
Check (-not $overPosted) "резерв 7 шт. при свободных 6 отклонен: $msg"

$freeOut = $c.Документы.ск_РасходТоваров.СоздатьДокумент()
$freeOut.Дата = Get-Date
$freeOut.Склад = $warehouse.Ссылка
$row = $freeOut.Товары.Добавить(); $row.Номенклатура = $item.Ссылка; $row.Количество = 7
$freeOutPosted = Post $freeOut
$msg = (Messages) -join ' | '
Check (-not $freeOutPosted) "расход 7 шт. без резерва (свободно 6) отклонен: $msg"

$outgo = $c.Документы.ск_РасходТоваров.СоздатьДокумент()
$outgo.Заполнить($reserve.Ссылка)
$outgo.Дата = Get-Date
Check ($outgo.Товары.Количество() -eq 1 -and $outgo.Товары.Получить(0).Количество -eq 4) "расход заполнен по резерву"
Check (Post $outgo) "расход 4 шт. по резерву проведен"

$q = $c.NewObject('Запрос')
$q.Текст = @"
ВЫБРАТЬ
	ЕСТЬNULL((ВЫБРАТЬ СУММА(О.КоличествоОстаток) ИЗ РегистрНакопления.ск_ТоварыНаСкладах.Остатки(, Номенклатура = &Номенклатура) КАК О), 0) КАК НаСкладе,
	ЕСТЬNULL((ВЫБРАТЬ СУММА(Р.КоличествоОстаток) ИЗ РегистрНакопления.ск_ТоварыВРезерве.Остатки(, Номенклатура = &Номенклатура) КАК Р), 0) КАК ВРезерве
"@
$q.УстановитьПараметр('Номенклатура', $item.Ссылка)
$sel = $q.Выполнить().Выбрать(); $sel.Следующий() | Out-Null
Check ($sel.НаСкладе -eq 6 -and $sel.ВРезерве -eq 0) "итог: на складе $($sel.НаСкладе), в резерве $($sel.ВРезерве)"

# Повторный резерв, чтобы отчет был не пустым, и проверка отчета.
$reserve2 = $c.Документы.ск_РезервТоваров.СоздатьДокумент()
$reserve2.Дата = Get-Date
$reserve2.Склад = $warehouse.Ссылка
$row = $reserve2.Товары.Добавить(); $row.Номенклатура = $item.Ссылка; $row.Количество = 2.5
Check (Post $reserve2) "резерв 2,5 шт. проведен"

$report = $c.Отчеты.ск_ТоварыВРезерве.Создать()
$doc = $c.NewObject('ТабличныйДокумент')
$report.СкомпоноватьРезультат($doc)
$reportFile = Join-Path $OutDir 'report.pdf'
$doc.Записать($reportFile, $c.ТипФайлаТабличногоДокумента.PDF)
Check ($doc.ВысотаТаблицы -gt 0) "отчет сформирован ($($doc.ВысотаТаблицы) строк): $reportFile"

# Этикетки.
$arr = $c.NewObject('Массив'); $arr.Добавить($item.Ссылка)
$labels = $c.ск_ПечатьЭтикетокВызовСервера.ЭтикеткиНоменклатуры($arr)
$labelFile = Join-Path $OutDir 'labels.pdf'
$labels.Записать($labelFile, $c.ТипФайлаТабличногоДокумента.PDF)
Check ($labels.ВысотаТаблицы -eq 4 -and $labels.Рисунки.Количество() -eq 1) "этикетка: $($labels.ВысотаТаблицы) строки, рисунков $($labels.Рисунки.Количество()): $labelFile"

$docs = $c.NewObject('Массив'); $docs.Добавить($income.Ссылка)
$labels2 = $c.ск_ПечатьЭтикетокВызовСервера.ЭтикеткиПоПриходам($docs)
Check ($labels2.Рисунки.Количество() -eq 10) "этикетки по приходу: $($labels2.Рисунки.Количество()) шт."

if ($failed) { Write-Host "Провалено проверок: $failed"; exit 1 } else { Write-Host 'Все проверки пройдены' }
