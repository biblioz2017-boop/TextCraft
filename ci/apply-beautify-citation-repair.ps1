$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}
function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}
function Normalize-Newlines([string]$Text, [string]$Nl) {
    return $Text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", $Nl)
}

$beautifyPath = 'GenerateUserControl.Beautify.cs'
$beautify = Read-Utf8Text $beautifyPath
$nl = if ($beautify.Contains("`r`n")) { "`r`n" } else { "`n" }

if (-not $beautify.Contains('NormalizeProtectedCitationKey(')) {
    $startMarker = '        private static List<string> FindMissingProtectedCitations(string source, string rewritten)'
    $start = $beautify.IndexOf($startMarker, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw 'Could not locate FindMissingProtectedCitations for citation normalization patch.'
    }

    $classTailMarker = $nl + '    }' + $nl + '}'
    $classTail = $beautify.LastIndexOf($classTailMarker, [System.StringComparison]::Ordinal)
    if ($classTail -lt 0 -or $classTail -le $start) {
        throw 'Could not locate GenerateUserControl class tail for citation normalization patch.'
    }

    $replacement = @'
        private static MatchCollection ExtractProtectedCitationMatches(string text)
        {
            return Regex.Matches(
                text ?? string.Empty,
                @"\[(?:S\d+|[^\]\r\n]{0,240}\.pdf[^\]\r\n]{0,120})\]",
                RegexOptions.IgnoreCase
            );
        }

        private static string NormalizeProtectedCitationKey(string citation)
        {
            if (string.IsNullOrWhiteSpace(citation))
                return string.Empty;

            string value = citation.Trim();
            if (value.Length >= 2 && value[0] == '[' && value[value.Length - 1] == ']')
                value = value.Substring(1, value.Length - 2);

            value = value.Replace('\u00A0', ' ');
            value = Regex.Replace(value, @"\s+", " ").Trim();

            Match sourceMarker = Regex.Match(value, @"^S(\d+)$", RegexOptions.IgnoreCase);
            int markerNumber;
            if (sourceMarker.Success && int.TryParse(sourceMarker.Groups[1].Value, out markerNumber))
                return "S:" + markerNumber;

            Match pdf = Regex.Match(
                value,
                @"^(.*?\.pdf)\s*,?\s*(?:(?:\u0441|\u0441\u0442\u0440|c|p)\.?\s*)?(\d+)\s*$",
                RegexOptions.IgnoreCase
            );
            if (pdf.Success)
            {
                string name = pdf.Groups[1].Value
                    .Normalize(NormalizationForm.FormC)
                    .ToLowerInvariant();
                name = Regex.Replace(name, @"[^\p{L}\p{Nd}]+", string.Empty);

                int pageNumber;
                if (int.TryParse(pdf.Groups[2].Value, out pageNumber))
                    return "PDF:" + name + "|P:" + pageNumber;
            }

            value = value.Normalize(NormalizationForm.FormC).ToLowerInvariant();
            return Regex.Replace(value, @"\s+", " ");
        }

        private static List<string> FindMissingProtectedCitations(string source, string rewritten)
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(source))
                return missing;

            var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in ExtractProtectedCitationMatches(rewritten))
            {
                string key = NormalizeProtectedCitationKey(match.Value);
                if (!string.IsNullOrWhiteSpace(key))
                    present.Add(key);
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in ExtractProtectedCitationMatches(source))
            {
                string citation = match.Value;
                string key = NormalizeProtectedCitationKey(citation);
                if (string.IsNullOrWhiteSpace(key) || !seen.Add(key))
                    continue;
                if (!present.Contains(key))
                    missing.Add(citation);
            }

            return missing;
        }

        private async Task<string> RepairMissingProtectedCitationsAsync(
            ChatClient client,
            string source,
            string draft,
            List<string> missingCitations,
            CancellationToken cancellationToken
        )
        {
            if (client == null || missingCitations == null || missingCitations.Count == 0)
                return draft ?? string.Empty;

            StringBuilder missingBlock = new StringBuilder();
            foreach (string citation in missingCitations)
                missingBlock.AppendLine(citation);

            string systemPrompt =
                "You are a precision citation-preservation editor. The DRAFT is already written. " +
                "Make the smallest possible edits needed to restore every MISSING PROTECTED CITATION next to the same claim it supported in SOURCE. " +
                "Do not add new facts, sources, bibliography entries, interpretations, or external knowledge. " +
                "Do not delete, merge, or reorder factual claims. Preserve every citation already present in DRAFT. " +
                "Return only the repaired draft, with no commentary.";

            string userPrompt =
                "MISSING PROTECTED CITATIONS:\n" + missingBlock.ToString() +
                "\nSOURCE:\n<<<SOURCE>>>\n" + (source ?? string.Empty) +
                "\n<<<END SOURCE>>>\n\nDRAFT:\n<<<DRAFT>>>\n" + (draft ?? string.Empty) +
                "\n<<<END DRAFT>>>";

            var answer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(userPrompt)
                },
                new ChatCompletionOptions { Temperature = 0.01f },
                cancellationToken
            );
            if (answer == null)
                return draft ?? string.Empty;

            string repaired = await CollectBeautifyStreamAsync(answer, cancellationToken);
            return string.IsNullOrWhiteSpace(repaired) ? (draft ?? string.Empty) : repaired;
        }
'@
    $replacement = Normalize-Newlines $replacement $nl
    $beautify = $beautify.Substring(0, $start) + $replacement + $beautify.Substring($classTail)
    Write-Utf8Text $beautifyPath $beautify
}

$ragPath = 'GenerateUserControl.BeautifyRag.cs'
$rag = Read-Utf8Text $ragPath
$ragNl = if ($rag.Contains("`r`n")) { "`r`n" } else { "`n" }

if (-not $rag.Contains('lostDuringRepair')) {
    # Use stable semantic anchors instead of matching the entire formatted block.
    $startMarker = '                List<string> missingCitations = FindMissingProtectedCitations('
    $endMarker = '                ReplaceBeautifyResponse(responseStart, grounded);'
    $start = $rag.IndexOf($startMarker, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw 'Could not locate citation-validation start marker in Beautify RAG.'
    }
    $end = $rag.IndexOf($endMarker, $start, [System.StringComparison]::Ordinal)
    if ($end -lt 0) {
        throw 'Could not locate citation-validation end marker in Beautify RAG.'
    }

    $new = @'
                List<string> missingCitations = FindMissingProtectedCitations(
                    boundedSource,
                    grounded
                );
                if (missingCitations.Count > 0)
                {
                    string repaired = await RepairMissingProtectedCitationsAsync(
                        client,
                        boundedSource,
                        grounded,
                        missingCitations,
                        cancellationToken
                    );

                    if (!string.IsNullOrWhiteSpace(repaired))
                    {
                        List<string> lostDuringRepair = FindMissingProtectedCitations(
                            grounded,
                            repaired
                        );
                        if (lostDuringRepair.Count == 0)
                            grounded = repaired;
                    }

                    missingCitations = FindMissingProtectedCitations(
                        boundedSource,
                        grounded
                    );
                }
                if (missingCitations.Count > 0)
                {
                    throw new InvalidOperationException(
                        "При переработке потеряны исходные ссылки: " +
                        string.Join(", ", missingCitations) + ". Выделенный текст Word не изменен."
                    );
                }

'@
    $new = Normalize-Newlines $new $ragNl
    $rag = $rag.Substring(0, $start) + $new + $rag.Substring($end)
    Write-Utf8Text $ragPath $rag
}

$verifyBeautify = Read-Utf8Text $beautifyPath
foreach ($marker in @(
    'NormalizeProtectedCitationKey',
    'ExtractProtectedCitationMatches',
    'RepairMissingProtectedCitationsAsync'
)) {
    if (-not $verifyBeautify.Contains($marker)) {
        throw ('Missing Beautify citation repair marker: ' + $marker)
    }
}

$verifyRag = Read-Utf8Text $ragPath
foreach ($marker in @(
    'lostDuringRepair',
    'RepairMissingProtectedCitationsAsync(',
    'ReplaceBeautifyResponse(responseStart, grounded);'
)) {
    if (-not $verifyRag.Contains($marker)) {
        throw ('Missing Beautify RAG citation repair marker: ' + $marker)
    }
}

Write-Host 'Beautify citation repair patch applied using stable range markers.'
