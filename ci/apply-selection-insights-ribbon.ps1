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

if (-not $text.Contains('ExplainSelectionButton_Click')) {
    $text = $text.Replace(
        '            this.SynonymsButton = this.Factory.CreateRibbonButton();',
        "            this.SynonymsButton = this.Factory.CreateRibbonButton();`r`n            this.ExplainSelectionButton = this.Factory.CreateRibbonButton();`r`n            this.ConclusionsSelectionButton = this.Factory.CreateRibbonButton();"
    )
    $text = $text.Replace(
        '            this.ToolsGroup.Items.Add(this.SynonymsButton);',
        "            this.ToolsGroup.Items.Add(this.SynonymsButton);`r`n            this.ToolsGroup.Items.Add(this.ExplainSelectionButton);`r`n            this.ToolsGroup.Items.Add(this.ConclusionsSelectionButton);"
    )
    $anchor = '            this.GrammarButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;'
    $insert = @"
            this.ExplainSelectionButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExplainSelectionButton.Image = global::TextForge.Properties.Resources.information_high_contrast;
            this.ExplainSelectionButton.Label = "Объяснить";
            this.ExplainSelectionButton.Name = "ExplainSelectionButton";
            this.ExplainSelectionButton.ShowImage = true;
            this.ExplainSelectionButton.SuperTip = "Объяснить выделенный фрагмент на основе его содержания, без добавления внешних фактов.";
            this.ExplainSelectionButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExplainSelectionButton_Click);

            this.ConclusionsSelectionButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ConclusionsSelectionButton.Image = global::TextForge.Properties.Resources.memo_high_contrast;
            this.ConclusionsSelectionButton.Label = "Выводы";
            this.ConclusionsSelectionButton.Name = "ConclusionsSelectionButton";
            this.ConclusionsSelectionButton.ShowImage = true;
            this.ConclusionsSelectionButton.SuperTip = "Сформулировать выводы только по выделенному фрагменту и вставить их после него.";
            this.ConclusionsSelectionButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ConclusionsSelectionButton_Click);

$anchor
"@
    if (-not $text.Contains($anchor)) { throw 'Grammar anchor not found.' }
    $text = $text.Replace($anchor, $insert)
    $text = $text.Replace(
        '        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;',
        "        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExplainSelectionButton;`r`n        internal Microsoft.Office.Tools.Ribbon.RibbonButton ConclusionsSelectionButton;`r`n        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;"
    )
    Write-Utf8Text $path $text
}

$verify = Read-Utf8Text $path
foreach ($marker in @(
    'this.ExplainSelectionButton = this.Factory.CreateRibbonButton();',
    'this.ConclusionsSelectionButton = this.Factory.CreateRibbonButton();',
    'this.ExplainSelectionButton_Click',
    'this.ConclusionsSelectionButton_Click'
)) {
    if (-not $verify.Contains($marker)) { throw ('Missing selection insight ribbon marker: ' + $marker) }
}

Write-Host 'Explain and conclusions ribbon buttons are wired.'
