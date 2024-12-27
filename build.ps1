# 设置 PowerShell 输出编码为 UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 获取当前的 git tag
$tag = git describe --tags --abbrev=0 2>$null

# 如果没有获取到 tag，则使用默认 tag v1.0.0
if (-not $tag) {
    $tag = "v1.0.0"
}

# 获取当前的 commit hash（前8位）
$hash = git rev-parse --short=8 HEAD

# 生成版本号
$version = "$tag.$hash"

# 获取当前日期，格式为 年.月.日 时:分:秒
$date = Get-Date -Format "yyyy.MM.dd HH:mm:ss"

# 读取 main.py 文件内容，并指定编码（使用UTF-8）
$mainFilePath = "main.py"
$fileContent = Get-Content $mainFilePath -Raw -Encoding UTF8

# 替换 VERSION 和 DATE 占位符
$fileContent = $fileContent -replace '@VERSION@', $version
$fileContent = $fileContent -replace '@DATE@', $date

# 将替换后的内容写入新的 RandQues.py 文件，使用 UTF-8 编码
$randQuesFilePath = "RandQues.py"
Set-Content $randQuesFilePath -Value $fileContent -Encoding UTF8

# 执行 pyinstaller 命令打包新的 RandQues.py 应用
$pyInstallerCommand = "pyinstaller --onefile --noconsole --icon=static/logo.ico $randQuesFilePath"

# 执行 pyinstaller 打包
Invoke-Expression $pyInstallerCommand

# 删除生成的 RandQues.py 文件
Remove-Item $randQuesFilePath  -Force -ErrorAction SilentlyContinue
Remove-Item -Path "./RandQues.spec" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "./build" -Recurse -Force -ErrorAction SilentlyContinue

# 输出版本号和日期
Write-Output "VERSION: $version"
Write-Output "DATE: $date"
Write-Output "PyInstaller SUCCESS"
