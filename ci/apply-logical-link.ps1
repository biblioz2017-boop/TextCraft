$ErrorActionPreference = 'Stop'

# This build-time patch MUST remain ASCII-only. Windows PowerShell 5.1 reads
# UTF-8 files without BOM through the active ANSI code page and would corrupt
# Cyrillic literals before they are written into Forge.Designer.cs.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

Write-Host 'Applying two-fragment logical-link ribbon button...'
$path = 'Forge.Designer.cs'
$designer = Read-Utf8Text $path

if (-not $designer.Contains('LogicalLinkButton')) {
    $fieldAnchor = '        private System.ComponentModel.IContainer components = null;'
    if (-not $designer.Contains($fieldAnchor)) {
        throw 'Could not locate Forge designer field anchor.'
    }
    $designer = $designer.Replace(
        $fieldAnchor,
        $fieldAnchor + "`r`n        private Microsoft.Office.Tools.Ribbon.RibbonButton LogicalLinkButton;"
    )

    $createAnchor = '            this.ContinueButton = this.Factory.CreateRibbonButton();'
    if (-not $designer.Contains($createAnchor)) {
        throw 'Could not locate ContinueButton creation anchor.'
    }
    $designer = $designer.Replace(
        $createAnchor,
        $createAnchor + "`r`n            this.LogicalLinkButton = this.Factory.CreateRibbonButton();"
    )

    $groupAnchor = '            this.ToolsGroup.Items.Add(this.ContinueButton);'
    if (-not $designer.Contains($groupAnchor)) {
        throw 'Could not locate tools group ContinueButton anchor.'
    }
    $designer = $designer.Replace(
        $groupAnchor,
        $groupAnchor + "`r`n            this.ToolsGroup.Items.Add(this.LogicalLinkButton);"
    )

    $generateAnchor = '            this.GenerateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;'
    if (-not $designer.Contains($generateAnchor)) {
        throw 'Could not locate GenerateButton configuration anchor.'
    }

    # Keep every character in this here-string ASCII. C# resolves the \uXXXX
    # escapes at compile time, so the Word ribbon receives proper Cyrillic text
    # regardless of the Windows PowerShell 5.1 system code page.
    $logicalBlock = @'
            this.LogicalLinkButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.LogicalLinkButton.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            this.LogicalLinkButton.Label = "\u0421\u0432\u044F\u0437\u0430\u0442\u044C";
            this.LogicalLinkButton.Name = "LogicalLinkButton";
            this.LogicalLinkButton.ShowImage = true;
            this.LogicalLinkButton.ScreenTip = "\u0421\u0432\u044F\u0437\u0430\u0442\u044C \u0434\u0432\u0430 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442\u0430";
            this.LogicalLinkButton.SuperTip = "\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u043F\u0435\u0440\u0432\u044B\u0439 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \u043A\u043D\u043E\u043F\u043A\u0443. \u0417\u0430\u0442\u0435\u043C \u0432\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0432\u0442\u043E\u0440\u043E\u0439 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \u0435\u0451 \u0435\u0449\u0451 \u0440\u0430\u0437. \u041D\u0435Z\u043D\u0430\u0439\u043A\u0430 \u0432\u0441\u0442\u0430\u0432\u0438\u0442 \u043A\u043E\u0440\u043E\u0442\u043A\u0438\u0439 \u043B\u043E\u0433\u0438\u0447\u0435\u0441\u043A\u0438\u0439 \u043F\u0435\u0440\u0435\u0445\u043E\u0434 \u043F\u0435\u0440\u0435\u0434 \u0432\u0442\u043E\u0440\u044B\u043C \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442\u043E\u043C, \u043D\u0435 \u043C\u0435\u043D\u044F\u044F \u0438\u0441\u0445\u043E\u0434\u043D\u044B\u0439 \u0442\u0435\u043A\u0441\u0442.";
            this.LogicalLinkButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogicalLinkButton_Click);

'@
    $designer = $designer.Replace($generateAnchor, $logicalBlock + $generateAnchor)
}

Write-Utf8Text $path $designer
Write-Host 'Logical-link button patch applied.'
