using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;

namespace TextForge
{
    public partial class Forge
    {
        private static RibbonButton _stopModelButton;
        private RibbonButton _unloadModelButton;
        private bool _modelControlsInitialized;

        private void InitializeModelControlButtons()
        {
            if (_modelControlsInitialized)
                return;

            _modelControlsInitialized = true;
            _optionsBox = this.OptionsGroup;
            _stopModelButton = this.CancelButton;

            this.OptionsGroup.Visible = true;
            this.CancelButton.Visible = true;
            this.CancelButton.Enabled = false;
            this.CancelButton.ScreenTip = "Остановить генерацию";
            this.CancelButton.SuperTip =
                "Прервать текущую генерацию ответа. Модель останется загруженной в памяти.";

            _unloadModelButton = this.Factory.CreateRibbonButton();
            _unloadModelButton.Name = "UnloadModelButton";
            _unloadModelButton.Label = "Выгрузить";
            _unloadModelButton.ScreenTip = "Выгрузить модель из памяти";
            _unloadModelButton.SuperTip =
                "Остановить текущую генерацию и выгрузить выбранную языковую модель из RAM/VRAM Ollama. " +
                "Следующий запрос автоматически загрузит модель снова.";
            _unloadModelButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            _unloadModelButton.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            _unloadModelButton.ShowImage = true;
            _unloadModelButton.Click += UnloadModelButton_Click;

            this.OptionsGroup.Items.Add(_unloadModelButton);
        }

        public static void SetModelActivity(bool active, string operation)
        {
            CancelButtonVisibility(active);
            if (active && !string.IsNullOrWhiteSpace(operation))
                SetStatus("◌ " + operation.Trim());
        }

        private async void UnloadModelButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (_unloadModelButton == null || !_unloadModelButton.Enabled)
                return;

            string model = ThisAddIn.Model;
            if (string.IsNullOrWhiteSpace(model))
                return;

            try
            {
                _unloadModelButton.Enabled = false;

                // Stop any active request first. Ollama releases the generation model
                // only after the request is no longer using it.
                try
                {
                    ThisAddIn.CancellationTokenSource.Cancel();
                }
                catch
                {
                }

                ThisAddIn.CancellationTokenSource = new CancellationTokenSource();
                CancelButtonVisibility(false);
                SetStatus("◌ Выгружаю…");

                await Task.Run(() => UnloadOllamaModel(model));
                SetStatus("● Модель выгружена");
            }
            catch (Exception ex)
            {
                SetStatus("● Ошибка выгрузки");
                CommonUtils.DisplayError(ex);
            }
            finally
            {
                if (_unloadModelButton != null)
                    _unloadModelButton.Enabled = true;
            }
        }

        private static void UnloadOllamaModel(string model)
        {
            Uri openAiEndpoint;
            if (!Uri.TryCreate(ThisAddIn.OpenAIEndpoint, UriKind.Absolute, out openAiEndpoint))
                throw new InvalidOperationException("Некорректный адрес Ollama/OpenAI endpoint.");

            UriBuilder nativeOllamaEndpoint = new UriBuilder(
                openAiEndpoint.Scheme,
                openAiEndpoint.Host,
                openAiEndpoint.Port,
                "/api/generate"
            );

            string json =
                "{\"model\":\"" + EscapeJsonString(model) +
                "\",\"prompt\":\"\",\"stream\":false,\"keep_alive\":0}";
            byte[] payload = Encoding.UTF8.GetBytes(json);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(nativeOllamaEndpoint.Uri);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.Timeout = 15000;
            request.ReadWriteTimeout = 15000;
            request.ContentLength = payload.Length;

            try
            {
                using (Stream requestStream = request.GetRequestStream())
                    requestStream.Write(payload, 0, payload.Length);

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    int statusCode = (int)response.StatusCode;
                    if (statusCode < 200 || statusCode >= 300)
                        throw new InvalidOperationException(
                            "Ollama не подтвердила выгрузку модели. HTTP " + statusCode + "."
                        );
                }
            }
            catch (WebException ex)
            {
                string details = string.Empty;
                HttpWebResponse response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    try
                    {
                        using (Stream stream = response.GetResponseStream())
                        using (StreamReader reader = new StreamReader(stream ?? Stream.Null))
                            details = reader.ReadToEnd();
                    }
                    catch
                    {
                    }
                }

                string message =
                    "Не удалось выгрузить модель через Ollama " + nativeOllamaEndpoint.Uri + ".";
                if (!string.IsNullOrWhiteSpace(details))
                    message += Environment.NewLine + details;

                throw new InvalidOperationException(message, ex);
            }
        }

        private static string EscapeJsonString(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }
    }
}
