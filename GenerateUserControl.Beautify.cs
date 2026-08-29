using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;

namespace TextForge
{
    public partial class GenerateUserControl
    {
        private Button _beautifyButton;
        private ContextMenuStrip _beautifyMenu;
        private bool _beautifyBusy;

        private sealed class BeautifyPreset
        {
            public string Name { get; private set; }
            public string Instruction { get; private set; }

            public BeautifyPreset(string name, string instruction)
            {
                Name = name;
                Instruction = instruction;
            }
        }

        private void InitializeBeautifyButton()
        {
            if (_beautifyButton != null)
                return;

            _beautifyButton = new Button
            {
                Text = "Сделать красиво",
                AutoSize = true,
                Height = 28
            };

            _beautifyMenu = new ContextMenuStrip();
            foreach (BeautifyPreset preset in GetBeautifyPresets())
            {
                var item = new ToolStripMenuItem(preset.Name)
                {
                    Tag = preset
                };
                item.Click += BeautifyPresetMenuItem_Click;
                _beautifyMenu.Items.Add(item);
            }

            _beautifyButton.Click += (s, e) =>
            {
                if (_beautifyBusy || _beautifyMenu == null || _beautifyButton == null)
                    return;

                _beautifyMenu.Show(
                    _beautifyButton,
                    new Point(0, _beautifyButton.Height)
                );
            };
        }

        private static List<BeautifyPreset> GetBeautifyPresets()
        {
            return new List<BeautifyPreset>
            {
                new BeautifyPreset(
                    "Научная статья",
                    "Преобразуй материал в цельный черновик научной статьи. Выстрой логическую структуру от постановки проблемы к основным данным, их интерпретации и выводам. Используй подзаголовки только когда они действительно помогают. Не придумывай разделы 'Материалы и методы' или новые результаты, если соответствующей информации в исходном материале нет."
                ),
                new BeautifyPreset(
                    "Раздел диссертации",
                    "Преобразуй материал в связный раздел диссертации: строгий научный стиль, последовательная аргументация, логические переходы между абзацами, умеренное использование подзаголовков. Удали разговорные формулировки и механические повторы, но не меняй научное содержание."
                ),
                new BeautifyPreset(
                    "Литературный обзор",
                    "Преобразуй материал в литературный обзор. Синтезируй близкие результаты, сопоставляй позиции источников, отдельно показывай противоречия, ограничения и пробелы. Не пересказывай источники по одному, если материал позволяет тематический синтез."
                ),
                new BeautifyPreset(
                    "Реферат",
                    "Преобразуй материал в академический реферат с введением, основной частью и заключением. Сохрани нейтральный научный тон и фактическую насыщенность. Не добавляй сведения, которых нет в исходном материале."
                ),
                new BeautifyPreset(
                    "Введение",
                    "Преобразуй материал во введение научной работы: обозначь проблему, ее актуальность только в пределах имеющихся данных, состояние вопроса и логический переход к дальнейшему изложению. Не придумывай цель, задачи, новизну или практическую значимость, если они явно не следуют из исходного материала."
                ),
                new BeautifyPreset(
                    "Обсуждение результатов",
                    "Преобразуй материал в раздел обсуждения результатов. Сопоставь имеющиеся результаты и интерпретации, подчеркни согласия и расхождения, ограничения и осторожные выводы. Не добавляй внешние объяснения или механизмы, отсутствующие в материале."
                ),
                new BeautifyPreset(
                    "Заключение / выводы",
                    "Преобразуй материал в компактное научное заключение. Выдели только выводы, которые действительно поддерживаются исходным текстом. Не вводи новые факты, причинные связи или рекомендации."
                ),
                new BeautifyPreset(
                    "Аннотация",
                    "Сожми материал до самостоятельной научной аннотации: предмет, ключевые положения, основные результаты или выводы и значение работы только в той мере, в какой это присутствует в исходном материале. Избегай ссылок на структуру текста вроде 'ниже рассмотрено'."
                ),
                new BeautifyPreset(
                    "Тезисы доклада",
                    "Преобразуй материал в краткие научные тезисы для доклада: проблема, ключевые положения, основные результаты и вывод. Сохрани информативность и убери второстепенные детали, не добавляя новых данных."
                )
            };
        }

        private async void BeautifyPresetMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            BeautifyPreset preset = item == null ? null : item.Tag as BeautifyPreset;
            if (preset == null)
                return;

            await BeautifyLastResponseAsync(preset);
        }

        private async Task BeautifyLastResponseAsync(BeautifyPreset preset)
        {
            if (_beautifyBusy)
                return;

            string source = _lastResponseMarkdown ?? string.Empty;
            if (string.IsNullOrWhiteSpace(source))
            {
                MessageBox.Show(
                    "Сначала получите материал в окне ответа, затем выберите формат в меню «Сделать красиво».",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            if (ModelProperties.IsImageModel(ThisAddIn.Model))
            {
                MessageBox.Show(
                    "Для редакторского прохода выберите языковую модель.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            string previousFinalResponse = _lastResponseMarkdown;
            string previousTemplateName = _lastTemplateName;

            try
            {
                _beautifyBusy = true;
                if (_beautifyButton != null)
                    _beautifyButton.Enabled = false;
                GenerateButton.Enabled = false;
                _insertButton.Enabled = false;
                _copyButton.Enabled = false;

                if (ThisAddIn.CancellationTokenSource == null ||
                    ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                {
                    try
                    {
                        ThisAddIn.CancellationTokenSource?.Dispose();
                    }
                    catch
                    {
                    }
                    ThisAddIn.CancellationTokenSource = new CancellationTokenSource();
                }

                int sourceTokens = Math.Max(2400, (int)(ThisAddIn.ContextLength * 0.62));
                string boundedSource = CommonUtils.SubstringTokens(source, sourceTokens);

                string originalQuestion = string.Empty;
                if (_conversationTurns.Count > 0)
                    originalQuestion = _conversationTurns[_conversationTurns.Count - 1].Question ?? string.Empty;

                string systemPrompt =
                    "Ты научный редактор. Твоя задача — только переработать предоставленный материал в выбранный жанр. " +
                    "Не используй внешние знания, не запускай поиск и не добавляй факты, которых нет в исходном тексте. " +
                    "Сохраняй без изменения числа, единицы измерения, формулы, названия веществ и методов, имена, даты, DOI, URL и цитаты. " +
                    "Все имеющиеся ссылки вида [S1], [S2] или [имя.pdf, с. N] сохраняй дословно и не переноси к утверждениям, которых они не подтверждали в исходном материале. " +
                    "Не выдумывай новые ссылки и библиографические данные. Если исходный материал содержит неопределенность или противоречие, сохрани ее. " +
                    "Можно улучшать композицию, связность, академический стиль, порядок абзацев и устранять повторы. Верни только готовый переработанный текст без комментариев о выполненной работе.";

                string userPrompt =
                    "Формат: " + preset.Name + "\n" +
                    "Инструкция: " + preset.Instruction + "\n" +
                    (string.IsNullOrWhiteSpace(originalQuestion)
                        ? string.Empty
                        : "Исходный вопрос/тема: " + originalQuestion + "\n") +
                    "\nМатериал для переработки:\n<<<SOURCE>>>\n" +
                    boundedSource +
                    "\n<<<END SOURCE>>>";

                ChatClient client = new ChatClient(
                    ThisAddIn.Model,
                    new ApiKeyCredential(ThisAddIn.ApiKey),
                    ThisAddIn.ClientOptions
                );

                var messages = new List<ChatMessage>
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(userPrompt)
                };

                Forge.SetModelActivity(true, "Оформляю: " + preset.Name + "…");

                _responseTextBox.AppendText(
                    Environment.NewLine + Environment.NewLine +
                    "НеZнайка [Сделать красиво → " + preset.Name + "]:" +
                    Environment.NewLine + Environment.NewLine
                );
                ScrollResponseToEnd();

                var answer = client.CompleteChatStreamingAsync(
                    messages,
                    new ChatCompletionOptions { Temperature = 0.12f },
                    ThisAddIn.CancellationTokenSource.Token
                );
                if (answer == null)
                    throw new InvalidOperationException("Модель не вернула редакторский ответ.");

                string rewritten = await StreamAnswerToPane(answer);
                if (string.IsNullOrWhiteSpace(rewritten))
                    throw new InvalidOperationException("Редакторский ответ модели оказался пустым.");

                List<string> missingCitations = FindMissingProtectedCitations(
                    boundedSource,
                    rewritten
                );
                if (missingCitations.Count > 0)
                {
                    _responseTextBox.AppendText(
                        Environment.NewLine + Environment.NewLine +
                        "[НеZнайка: результат не принят как готовый — модель потеряла ссылки: " +
                        string.Join(", ", missingCitations) + "]"
                    );
                    ScrollResponseToEnd();
                    throw new InvalidOperationException(
                        "Редакторский проход потерял одну или несколько исходных ссылок. Исходный материал сохранен как текущий готовый ответ."
                    );
                }

                _lastResponseMarkdown = rewritten;
                _lastTemplateName = "Сделать красиво — " + preset.Name;
                _insertButton.Enabled = true;
                _copyButton.Enabled = true;
            }
            catch (OperationCanceledException)
            {
                _lastResponseMarkdown = previousFinalResponse;
                _lastTemplateName = previousTemplateName;
            }
            catch (Exception ex)
            {
                _lastResponseMarkdown = previousFinalResponse;
                _lastTemplateName = previousTemplateName;
                CommonUtils.DisplayWarning(ex);
            }
            finally
            {
                try { Forge.SetModelActivity(false, null); } catch { }
                _beautifyBusy = false;
                GenerateButton.Enabled = true;
                if (_beautifyButton != null && !_beautifyButton.IsDisposed)
                    _beautifyButton.Enabled = true;
                _insertButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                _copyButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                ScrollResponseToEnd();
            }
        }

        private static List<string> FindMissingProtectedCitations(string source, string rewritten)
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(source))
                return missing;

            MatchCollection matches = Regex.Matches(
                source,
                @"\[(?:S\d+|[^\]\r\n]{0,160}\.pdf[^\]\r\n]{0,80})\]",
                RegexOptions.IgnoreCase
            );

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match match in matches)
            {
                string citation = match.Value;
                if (!seen.Add(citation))
                    continue;
                if ((rewritten ?? string.Empty).IndexOf(citation, StringComparison.Ordinal) < 0)
                    missing.Add(citation);
            }

            return missing;
        }
    }
}
