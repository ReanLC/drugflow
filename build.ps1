# 编译脚本：使用 Windows 自带的 .NET Framework 4.x 编译器，无需联网、无需安装任何 SDK
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src  = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
# 用 32 位 csc：目标程序是 32 位，这样引用程序集版本才匹配（Access/Jet 也多为 32 位）
$csc  = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw "找不到 C# 编译器 csc.exe（需要 .NET Framework 4.x）" }

if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

$files = Get-ChildItem -Path $src -Filter *.cs | ForEach-Object { $_.FullName }
$out = Join-Path $dist 'DrugFlow.exe'

$refs = @(
  '/r:System.dll',
  '/r:System.Core.dll',
  '/r:System.Data.dll',
  '/r:System.Drawing.dll',
  '/r:System.Windows.Forms.dll',
  '/r:System.Xml.dll'
)
$oleDb = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\System.Data.OleDb.dll'
$odbc  = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\System.Data.Odbc.dll'
if (Test-Path $oleDb) { $refs += '/r:System.Data.OleDb.dll' }
if (Test-Path $odbc)  { $refs += '/r:System.Data.Odbc.dll' }

$args = @('/nologo', '/target:winexe', '/platform:x86', '/optimize+', '/codepage:65001', "/out:$out") + $refs + $files
& $csc $args
if ($LASTEXITCODE -ne 0) { throw "编译失败，退出码 $LASTEXITCODE" }

Write-Host "编译完成：$out"
Get-Item $out | Select-Object Name, @{n='大小KB';e={[math]::Round($_.Length/1KB,1)}}, LastWriteTime | Format-List
