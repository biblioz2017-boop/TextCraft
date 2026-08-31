$ErrorActionPreference = 'Stop'

# ASCII-only build patch for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

$path = 'Forge.Designer.cs'
$text = Read-Utf8Text $path
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

if (-not $text.Contains('this.HelpButton = this.Factory.CreateRibbonButton();')) {
    $createMarker = '            this.AboutButton = this.Factory.CreateRibbonButton();'
    if (-not $text.Contains($createMarker)) { throw 'Help UI create marker not found.' }
    $createBlock = @(
        '            this.HelpButton = this.Factory.CreateRibbonButton();',
        $createMarker
    ) -join $nl
    $text = $text.Replace($createMarker, $createBlock)

    $itemsMarker = '            this.InfoGroup.Items.Add(this.AboutButton);'
    if (-not $text.Contains($itemsMarker)) { throw 'Help UI info group marker not found.' }
    $itemsBlock = @(
        '            this.InfoGroup.Items.Add(this.HelpButton);',
        $itemsMarker
    ) -join $nl
    $text = $text.Replace($itemsMarker, $itemsBlock)

    $configMarker = '            this.AboutButton.Image = global::TextForge.Properties.Resources.information_high_contrast;'
    if (-not $text.Contains($configMarker)) { throw 'Help UI config marker not found.' }
    $configBlock = @(
        '            this.HelpButton.Image = global::TextForge.Properties.Resources.information_high_contrast;',
        '            this.HelpButton.Label = "\u0420\u0443\u043A\u043E\u0432\u043E\u0434\u0441\u0442\u0432\u043E";',
        '            this.HelpButton.Name = "HelpButton";',
        '            this.HelpButton.ShowImage = true;',
        '            this.HelpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HelpButton_Click);',
        '',
        $configMarker
    ) -join $nl
    $text = $text.Replace($configMarker, $configBlock)

    $fieldMarker = '        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;'
    if (-not $text.Contains($fieldMarker)) { throw 'Help UI field marker not found.' }
    $fieldBlock = @(
        '        internal Microsoft.Office.Tools.Ribbon.RibbonButton HelpButton;',
        $fieldMarker
    ) -join $nl
    $text = $text.Replace($fieldMarker, $fieldBlock)

    $localizeMarker = '            this.DefaultCheckBox.Visible = false;'
    if (-not $text.Contains($localizeMarker)) { throw 'Help UI localization marker not found.' }
    $localizeBlock = @(
        $localizeMarker,
        '            this.HelpButton.Label = "\u0420\u0443\u043A\u043E\u0432\u043E\u0434\u0441\u0442\u0432\u043E";',
        '            this.HelpButton.SuperTip = "\u041E\u0442\u043A\u0440\u044B\u0442\u044C \u0432\u0441\u0442\u0440\u043E\u0435\u043D\u043D\u043E\u0435 \u043F\u043E\u0434\u0440\u043E\u0431\u043D\u043E\u0435 \u0440\u0443\u043A\u043E\u0432\u043E\u0434\u0441\u0442\u0432\u043E \u043F\u043E \u0440\u0430\u0431\u043E\u0442\u0435 \u0441 \u041D\u0435Z\u043D\u0430\u0439\u043A\u0430.";'
    ) -join $nl
    $text = $text.Replace($localizeMarker, $localizeBlock)

    Write-Utf8Text $path $text
}

$verify = Read-Utf8Text $path
foreach ($marker in @(
    'this.HelpButton = this.Factory.CreateRibbonButton();',
    'this.InfoGroup.Items.Add(this.HelpButton);',
    'this.HelpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HelpButton_Click);',
    'RibbonButton HelpButton;'
)) {
    if (-not $verify.Contains($marker)) {
        throw ('Missing Help UI marker: ' + $marker)
    }
}

Write-Host 'NeZnaika embedded help ribbon UI applied.'
