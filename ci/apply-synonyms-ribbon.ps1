$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

$path = 'Forge.Designer.cs'
$text = Read-Utf8Text $path

if (-not $text.Contains('SynonymsButton_Click')) {
    $pairs = @(
        @(
            '            this.KeywordsButton = this.Factory.CreateRibbonButton();',
            "            this.KeywordsButton = this.Factory.CreateRibbonButton();`r`n            this.SynonymsButton = this.Factory.CreateRibbonButton();"
        ),
        @(
            '            this.ToolsGroup.Items.Add(this.KeywordsButton);',
            "            this.ToolsGroup.Items.Add(this.KeywordsButton);`r`n            this.ToolsGroup.Items.Add(this.SynonymsButton);"
        ),
        @(
            '            this.GrammarButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;',
            "            this.SynonymsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;`r`n            this.SynonymsButton.Image = global::TextForge.Properties.Resources.face_with_monocle_high_contrast;`r`n            this.SynonymsButton.Label = \"\u0421\u0438\u043D\u043E\u043D\u0438\u043C\u044B\";`r`n            this.SynonymsButton.Name = \"SynonymsButton\";`r`n            this.SynonymsButton.ShowImage = true;`r`n            this.SynonymsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SynonymsButton_Click);`r`n`r`n            this.GrammarButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;"
        ),
        @(
            '        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;',
            "        internal Microsoft.Office.Tools.Ribbon.RibbonButton SynonymsButton;`r`n        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;"
        )
    )

    foreach ($pair in $pairs) {
        $old = $pair[0]
        $new = $pair[1]
        if (-not $text.Contains($old)) {
            throw ('Could not locate ribbon anchor: ' + $old)
        }
        $text = $text.Replace($old, $new)
    }

    $localizeAnchor = '            this.GrammarButton.Label = "\u0413\u0440\u0430\u043C\u043C\u0430\u0442\u0438\u043A\u0430";'
    if (-not $text.Contains($localizeAnchor)) {
        # The source normally contains literal Cyrillic. Insert before the stable GrammarButton.SuperTip line instead.
        $localizeAnchor = '            this.GrammarButton.SuperTip = '
        $pos = $text.IndexOf($localizeAnchor, [System.StringComparison]::Ordinal)
        if ($pos -lt 0) { throw 'Could not locate GrammarButton localization anchor.' }
        $insert = "            this.SynonymsButton.Label = \"\u0421\u0438\u043D\u043E\u043D\u0438\u043C\u044B\";`r`n            this.SynonymsButton.SuperTip = \"\u041F\u043E\u0434\u043E\u0431\u0440\u0430\u0442\u044C \u043A\u043E\u043D\u0442\u0435\u043A\u0441\u0442\u043D\u044B\u0435 \u0441\u0438\u043D\u043E\u043D\u0438\u043C\u044B \u0434\u043B\u044F \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u043D\u043E\u0433\u043E \u0441\u043B\u043E\u0432\u0430 \u0438\u043B\u0438 \u0444\u0440\u0430\u0437\u044B \u0441 \u043F\u043E\u043C\u043E\u0449\u044C\u044E \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0439 LLM.\";`r`n"
        $text = $text.Substring(0, $pos) + $insert + $text.Substring($pos)
    }

    Write-Utf8Text $path $text
}

$verify = Read-Utf8Text $path
foreach ($marker in @(
    'this.SynonymsButton = this.Factory.CreateRibbonButton();',
    'this.ToolsGroup.Items.Add(this.SynonymsButton);',
    'this.SynonymsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SynonymsButton_Click);',
    'RibbonButton SynonymsButton;'
)) {
    if (-not $verify.Contains($marker)) {
        throw ('Missing synonym ribbon marker: ' + $marker)
    }
}

Write-Host 'Contextual synonym ribbon button is wired.'
