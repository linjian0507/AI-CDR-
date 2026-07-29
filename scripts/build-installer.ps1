param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$installer = Join-Path $repo "installer"
$release = Join-Path $repo "release"
$payloadRoot = Join-Path $installer "payload\AIVectorHelper"
$payloadZip = Join-Path $installer "payload.zip"

New-Item -ItemType Directory -Force -Path $release | Out-Null

$forbiddenNames = @("config.json", "history.json", "svg-history.tsv")
$forbidden = Get-ChildItem -LiteralPath $payloadRoot -Recurse -Force -File |
    Where-Object { $_.Name -in $forbiddenNames -or $_.FullName -match "(?i)super.?svg" }
if ($forbidden) {
    throw "payload 包含禁止发布的用户配置、历史文件或 SuperSVG 文件：`n$($forbidden.FullName -join "`n")"
}

if (Test-Path -LiteralPath $payloadZip) {
    Remove-Item -LiteralPath (Resolve-Path -LiteralPath $payloadZip).Path -Force
}
Compress-Archive -Path (Join-Path $payloadRoot "*") `
    -DestinationPath $payloadZip `
    -CompressionLevel Optimal

dotnet build (Join-Path $installer "AIVectorInstaller.csproj") -c $Configuration

$built = Join-Path $installer "bin\$Configuration\net48\AIVectorInstaller.exe"
$target = Join-Path $release "AI矢量助手-安装程序-v2.3.3.exe"
Copy-Item -LiteralPath $built -Destination $target -Force

$zip = Join-Path $release "AI矢量助手-v2.3.3-安装包.zip"
if (Test-Path -LiteralPath $zip) {
    Remove-Item -LiteralPath (Resolve-Path -LiteralPath $zip).Path -Force
}
$readme = Join-Path $release "安装说明.txt"
@"
AI 矢量助手 v2.3.3 安装包

1. 关闭所有 CorelDRAW。
2. 右键 AI矢量助手-安装程序-v2.3.1.exe，选择“以管理员身份运行”。
3. 勾选要安装的 CorelDRAW 版本，点击“安装”。
4. 安装完成后重新启动 CorelDRAW。
5. 如需卸载，重新运行安装程序，勾选目标版本后点击“卸载选中”。

此发布包不包含 config.json、history.json、svg-history.tsv、API Key、生成历史和模型参数。
安装时不会覆盖目标 CDR 安装中已有的 config.json、svg-history.tsv 和 output。
"@ | Set-Content -LiteralPath $readme -Encoding UTF8
Compress-Archive -Path $target, $readme -DestinationPath $zip -CompressionLevel Optimal

Write-Host "安装程序：" $target
Write-Host "压缩包：" $zip
