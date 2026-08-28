param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$rootPath = (Resolve-Path -LiteralPath $Root).Path

$requiredFiles = @(
    'README.md',
    'docs/INDEX.md',
    'docs/ARCHITECTURE.md',
    'docs/DEVELOPMENT.md',
    'docs/DEPLOYMENT.md',
    'docs/ENVIRONMENT-INVENTORY.md',
    'docs/DATA-DICTIONARY.md',
    'docs/POWER-AUTOMATE.md',
    'docs/OPERATIONS-RUNBOOK.md',
    'docs/TROUBLESHOOTING.md',
    'docs/SECURITY-ACCESS.md',
    'docs/LGPD-DATA-GOVERNANCE.md',
    'docs/DISASTER-RECOVERY.md',
    'docs/OWNERS.md',
    'docs/CONTINUITY-HANDOVER.md',
    'docs/CONTINUITY-EVIDENCE.md',
    'docs/decisions/ADR-0001-SPFX-E-MICROSOFT-365.md',
    'docs/checklists/RELEASE.md',
    'docs/checklists/HANDOVER-VALIDATION.md'
)

$errors = [System.Collections.Generic.List[string]]::new()

foreach ($relativeFile in $requiredFiles) {
    $fullPath = Join-Path $rootPath $relativeFile
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        $errors.Add("Arquivo obrigatório ausente: $relativeFile")
        continue
    }

    $content = Get-Content -Raw -LiteralPath $fullPath
    if ([string]::IsNullOrWhiteSpace($content)) {
        $errors.Add("Arquivo vazio: $relativeFile")
    }
    elseif ($content -notmatch '(?m)^#\s+\S') {
        $errors.Add("Título H1 ausente: $relativeFile")
    }
}

$markdownFiles = @(
    Get-Item -LiteralPath (Join-Path $rootPath 'README.md')
    Get-ChildItem -LiteralPath (Join-Path $rootPath 'docs') -Filter '*.md' -File -Recurse
)
$linkPattern = '\[[^\]]+\]\((?<target>[^)]+)\)'
$checkedLinks = 0

foreach ($markdownFile in $markdownFiles) {
    $content = Get-Content -Raw -LiteralPath $markdownFile.FullName
    if ($null -eq $content) {
        $content = ''
    }
    foreach ($match in [regex]::Matches($content, $linkPattern)) {
        $target = $match.Groups['target'].Value.Trim()
        if ($target.StartsWith('<') -and $target.EndsWith('>')) {
            $target = $target.Substring(1, $target.Length - 2)
        }
        if ($target -match '^(https?://|mailto:|#)') {
            continue
        }

        $checkedLinks++
        $pathPart = ($target -split '#', 2)[0]
        if ([string]::IsNullOrWhiteSpace($pathPart)) {
            continue
        }

        $decodedPath = [System.Uri]::UnescapeDataString($pathPart)
        $resolvedTarget = Join-Path $markdownFile.DirectoryName $decodedPath
        if (-not (Test-Path -LiteralPath $resolvedTarget)) {
            $relativeSource = $markdownFile.FullName.Replace(($rootPath + [System.IO.Path]::DirectorySeparatorChar), '')
            $errors.Add("Link local quebrado em ${relativeSource}: $target")
        }
    }
}

Write-Host "Arquivos Markdown: $($markdownFiles.Count)"
Write-Host "Arquivos obrigatórios: $($requiredFiles.Count)"
Write-Host "Links locais verificados: $checkedLinks"

if ($errors.Count -gt 0) {
    $errors | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Host 'Documentação validada sem arquivos obrigatórios ausentes ou links locais quebrados.'
