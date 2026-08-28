$ErrorActionPreference = 'Stop'

$targetPath = 'GenerateUserControl.Science.cs'
$fragmentPath = 'ci/GenerateUserControl.Continue.fragment.cs.txt'
$marker = 'private System.Windows.Forms.Button _continueResponseButton;'

$targetFull = (Resolve-Path $targetPath).Path
$fragmentFull = (Resolve-Path $fragmentPath).Path
$target = [System.IO.File]::ReadAllText($targetFull, [System.Text.Encoding]::UTF8)

if (-not $target.Contains($marker)) {
    $fragment = [System.IO.File]::ReadAllText($fragmentFull, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText(
        $targetFull,
        $target.TrimEnd() + [Environment]::NewLine + $fragment,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

$loadAnchor = '            AddEvidencePanel();'
$loadCall = '            EnsureContinueResponseButton();'
$target = [System.IO.File]::ReadAllText($targetFull, [System.Text.Encoding]::UTF8)
if (-not $target.Contains($loadCall)) {
    if (-not $target.Contains($loadAnchor)) {
        throw 'Could not locate chat OnLoad anchor for continuation button.'
    }
    $target = $target.Replace(
        $loadAnchor,
        $loadAnchor + [Environment]::NewLine + $loadCall
    )
    [System.IO.File]::WriteAllText(
        $targetFull,
        $target,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

Write-Host 'Continuation response button prepared.'
