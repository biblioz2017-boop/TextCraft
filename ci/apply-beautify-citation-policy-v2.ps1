$ErrorActionPreference = 'Stop'

# ASCII-only build patch for Windows PowerShell 5.1.
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

if (-not $beautify.Contains('NormalizeProtectedCitationKeyV2(')) {
    $startMarker = '        private static List<string> FindMissingProtectedCitations(string source, string rewritten)'
    $start = $beautify.IndexOf($startMarker, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw 'Could not locate FindMissingProtectedCitations for citation policy v2.'
    }

    $classTailMarker = $nl + '    }' + $nl + '}'
    $classTail = $beautify.LastIndexOf($classTailMarker, [System.StringComparison]::Ordinal)
    if ($classTail -lt 0 -or $classTail -le $start) {
        throw 'Could not locate GenerateUserControl class tail for citation policy v2.'
    }

    $replacement = @'
        private static MatchCollection ExtractProtectedCitationMatchesV2(string text)
        {
            return Regex.Matches(
                text ?? string.Empty,
                @"\[(?:S\s*\d+|[^\]\r\n]{0,260}\.pdf[^\]\r\n]{0,140})\]",
                RegexOptions.IgnoreCase
            );
        }

        private static string NormalizeProtectedCitationKeyV2(string citation)
        {
            if (string.IsNullOrWhiteSpace(citation))
                return string.Empty;

            string value = citation.Trim();
            if (value.Length >= 2 && value[0] == '[' && value[value.Length - 1] == ']')
                value = value.Substring(1, value.Length - 2);

            value = value.Replace('\u00A0', ' ').Replace('\u202F', ' ');
            value = Regex.Replace(value, @"\s+", " ").Trim();

            Match sourceMarker = Regex.Match(value, @"^S\s*(\d+)$", RegexOptions.IgnoreCase);
            int markerNumber;
            if (sourceMarker.Success && int.TryParse(sourceMarker.Groups[1].Value, out markerNumber))
                return "S:" + markerNumber;

            Match pdf = Regex.Match(
                value,
                @"^(.*?\.pdf)\s*[,;]?\s*(?:(?:\u0441|\u0441\u0442\u0440|c|p|page)\.?\s*)?(\d+)(?:\s*[-\u2013\u2014]\s*(\d+))?\s*$",
                RegexOptions.IgnoreCase
            );

            if (pdf.Success)
            {
                string name = pdf.Groups[1].Value
                    .Normalize(NormalizationForm.FormKC)
                    .ToLowerInvariant();
                name = Regex.Replace(name, @"[^\p{L}\p{Nd}]+", string.Empty);

                int firstPage;
                int lastPage;
                if (int.TryParse(pdf.Groups[2].Value, out firstPage))
                {
                    string key = "PDF:" + name + "|P:" + firstPage;
                    if (pdf.Groups[3].Success &&
                        int.TryParse(pdf.Groups[3].Value, out lastPage) &&
                        lastPage != firstPage)
                    {
                        key += "-" + lastPage;
                    }
                    return key;
                }
            }

            value = value.Normalize(NormalizationForm.FormKC).ToLowerInvariant();
            return Regex.Replace(value, @"\s+", " ");
        }

        private static List<string> FindMissingProtectedCitations(string source, string rewritten)
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(source))
                return missing;

            var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in ExtractProtectedCitationMatchesV2(rewritten))
            {
                string key = NormalizeProtectedCitationKeyV2(match.Value);
                if (!string.IsNullOrWhiteSpace(key))
                    present.Add(key);
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in ExtractProtectedCitationMatchesV2(source))
            {
                string citation = match.Value;
                string key = NormalizeProtectedCitationKeyV2(citation);
                if (string.IsNullOrWhiteSpace(key) || !seen.Add(key))
                    continue;
                if (!present.Contains(key))
                    missing.Add(citation);
            }

            return missing;
        }

        private static string BuildCitationRepairHintsV2(string source, List<string> missingCitations)
        {
            var hints = new StringBuilder();
            if (string.IsNullOrWhiteSpace(source) || missingCitations == null)
                return string.Empty;

            int number = 1;
            foreach (string citation in missingCitations)
            {
                if (string.IsNullOrWhiteSpace(citation))
                    continue;

                int index = source.IndexOf(citation, StringComparison.Ordinal);
                int start = index >= 0 ? Math.Max(0, index - 320) : 0;
                int end = index >= 0
                    ? Math.Min(source.Length, index + citation.Length + 320)
                    : Math.Min(source.Length, 640);

                string context = source.Substring(start, Math.Max(0, end - start));
                context = Regex.Replace(context, @"\s+", " ").Trim();

                hints.Append(number++).Append(") ").AppendLine(citation);
                hints.Append("SOURCE CONTEXT: ").AppendLine(context);
                hints.AppendLine();
            }

            return hints.ToString();
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

            string hints = BuildCitationRepairHintsV2(source, missingCitations);
            string systemPrompt =
                "You are a precision citation-preservation editor. The DRAFT is already complete. " +
                "Restore only the missing protected citations shown in CITATION HINTS. " +
                "Place each citation next to the claim supported by its SOURCE CONTEXT. " +
                "Do not add facts, sources, bibliography entries, interpretations, headings, or external knowledge. " +
                "Do not delete or rewrite factual claims except for the smallest punctuation change needed to place a citation. " +
                "Preserve every citation already present in DRAFT. Return only the repaired draft.";

            string userPrompt =
                "CITATION HINTS:\n" + hints +
                "\nSOURCE:\n<<<SOURCE>>>\n" + (source ?? string.Empty) +
                "\n<<<END SOURCE>>>\n\nDRAFT:\n<<<DRAFT>>>\n" + (draft ?? string.Empty) +
                "\n<<<END DRAFT>>>";

            var answer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(userPrompt)
                },
                new ChatCompletionOptions { Temperature = 0.0f },
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

if (-not $rag.Contains('citationValidationWarningV2')) {
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
                        List<string> lostDuringRepairV2 = FindMissingProtectedCitations(
                            grounded,
                            repaired
                        );
                        if (lostDuringRepairV2.Count == 0)
                            grounded = repaired;
                    }

                    missingCitations = FindMissingProtectedCitations(
                        boundedSource,
                        grounded
                    );
                }

                bool citationValidationWarningV2 = missingCitations.Count > 0;

                ReplaceBeautifyResponse(responseStart, grounded);

                if (citationValidationWarningV2)
                {
                    _responseTextBox.AppendText(
                        Environment.NewLine + Environment.NewLine +
                        "\u26A0 NeZnaika: \u0447\u0430\u0441\u0442\u044C \u0438\u0441\u0445\u043E\u0434\u043D\u044B\u0445 \u0441\u0441\u044B\u043B\u043E\u043A \u043D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0430\u0432\u0442\u043E\u043C\u0430\u0442\u0438\u0447\u0435\u0441\u043A\u0438 \u0432\u043E\u0441\u0441\u0442\u0430\u043D\u043E\u0432\u0438\u0442\u044C. \u041F\u0440\u043E\u0432\u0435\u0440\u044C\u0442\u0435 \u0438\u0445 \u043F\u0435\u0440\u0435\u0434 \u0432\u0441\u0442\u0430\u0432\u043A\u043E\u0439 \u0432 Word: " +
                        string.Join(", ", missingCitations)
                    );
                    ScrollResponseToEnd();
                }

'@
    $new = Normalize-Newlines $new $ragNl
    $rag = $rag.Substring(0, $start) + $new + $rag.Substring($end + $endMarker.Length)
    Write-Utf8Text $ragPath $rag
}

$verifyBeautify = Read-Utf8Text $beautifyPath
foreach ($marker in @(
    'NormalizeProtectedCitationKeyV2',
    'BuildCitationRepairHintsV2',
    'RepairMissingProtectedCitationsAsync'
)) {
    if (-not $verifyBeautify.Contains($marker)) {
        throw ('Missing Beautify citation policy v2 marker: ' + $marker)
    }
}

$verifyRag = Read-Utf8Text $ragPath
foreach ($marker in @(
    'citationValidationWarningV2',
    'lostDuringRepairV2',
    'ReplaceBeautifyResponse(responseStart, grounded);'
)) {
    if (-not $verifyRag.Contains($marker)) {
        throw ('Missing Beautify RAG citation policy v2 marker: ' + $marker)
    }
}

Write-Host 'Beautify citation policy v2 applied.'
