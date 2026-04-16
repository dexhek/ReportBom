param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = if (Test-Path (Join-Path $scriptDirectory "ReportBom.AddIn")) {
    $scriptDirectory
}
elseif ((Split-Path -Leaf $scriptDirectory) -eq $Configuration -and (Split-Path -Leaf (Split-Path -Parent $scriptDirectory)) -eq "bin") {
    Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDirectory))
}
else {
    $scriptDirectory
}

$candidatePaths = @(
    (Join-Path $repoRoot "ReportBom.AddIn\\bin\\$Configuration\\ReportBom.AddIn.dll"),
    (Join-Path $scriptDirectory "ReportBom.AddIn.dll")
)

$addInDll = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
$regAsmPath = Join-Path $env:WINDIR "Microsoft.NET\\Framework64\\v4.0.30319\\RegAsm.exe"

if (-not $addInDll) {
    throw "Add-in non trovato. Cercato in: $($candidatePaths -join ', ')."
}

if (-not (Test-Path $regAsmPath)) {
    throw "RegAsm non trovato: $regAsmPath"
}

& $regAsmPath $addInDll /codebase
