using LLama;
using LLama.Common;
using LLama.Sampling;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace PDFControls
{
    /// <summary>
    /// Extracts normalized report dates with a small local GGUF model.
    /// The model is deliberately used only as a semantic supplement to Regex.
    /// </summary>
    public sealed class AiPdfTextExtractor : IDisposable
    {
        private const int ModelContextSize = 2048;
        private const int MaxSourceCharacters = 2400;
        private const int MaxOutputTokens = 64;
        private const int MaxOutputTokensWithProducts = 128;
        private readonly int _cpuPercent;
        private readonly LLamaWeights _weights;
        private readonly StatelessExecutor _executor;
        private readonly SemaphoreSlim _gate = new(1, 1);

        public AiPdfTextExtractor(string modelPath, int cpuPercent = 75)
        {
            if (!File.Exists(modelPath))
                throw new FileNotFoundException("AI model was not found.", modelPath);

            _cpuPercent = Math.Clamp(cpuPercent, 1, 100);
            var parameters = new ModelParams(modelPath)
            {
                ContextSize = ModelContextSize,
                GpuLayerCount = 0,
                Threads = GetCpuThreadLimit(),
                BatchThreads = GetCpuThreadLimit()
            };
            _weights = LLamaWeights.LoadFromFile(parameters);
            _executor = new StatelessExecutor(_weights, parameters);
        }

        public async Task<PdfTextComparer.PdfCodeInfo?> ExtractAsync(
            string text,
            IEnumerable<string>? contexts = null,
            CancellationToken token = default,
            bool includeProductCount = false)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            // Leave room for instructions and generated JSON inside the 2048-token KV cache.
            var source = text.Length > MaxSourceCharacters ? text.Substring(0, MaxSourceCharacters) : text;
            var prompt = BuildPrompt(source, contexts, includeProductCount);
            var output = new StringBuilder();
            var inference = new InferenceParams
            {
                MaxTokens = includeProductCount ? MaxOutputTokensWithProducts : MaxOutputTokens,
                AntiPrompts = new List<string> { "\nSOURCE:", "\nUSER:" },
                SamplingPipeline = new DefaultSamplingPipeline { Temperature = 0 }
            };

            await _gate.WaitAsync(token).ConfigureAwait(false);
            try
            {
                await foreach (var part in _executor.InferAsync(prompt, inference).WithCancellation(token).ConfigureAwait(false))
                    output.Append(part);
            }
            finally
            {
                _gate.Release();
            }

            return ParseResponse(output.ToString(), includeProductCount);
        }

        private static string BuildPrompt(string source, IEnumerable<string>? contexts, bool includeProductCount)
        {
            var configuredContexts = contexts == null ? string.Empty : string.Join("; ", contexts.Where(c => !string.IsNullOrWhiteSpace(c)));
            var keys = includeProductCount
                ? "latest_code, issue_date, expiration_date and products_count"
                : "latest_code, issue_date and expiration_date";
            var valueRule = includeProductCount
                ? " Values must be a code, ISO yyyy-MM-dd or null. products_count must be an integer (number of distinct products listed on the page) or null.\n"
                : " Values must be a code, ISO yyyy-MM-dd or null.\n";
            var example = includeProductCount
                ? "JSON: {\"issue_date\":\"2024-03-15\",\"expiration_date\":\"2027-03-15\",\"products_count\":null}\n"
                : "JSON: {\"issue_date\":\"2024-03-15\",\"expiration_date\":\"2027-03-15\"}\n";
            var countRule = includeProductCount
                ? " products_count must be the ACTUAL number of distinct products on this page, never a fixed or example number."
                : string.Empty;
            return "You extract data from technical evaluation reports. Return ONLY one JSON object with keys " +
                keys + "." +
                valueRule +
                countRule +
                "Issue/Rev Date means Revised, Reissued, Updated, Issued Date, Date of Revision. " +
                "Expiration Date means Expiration, Active through, Valid through, Valid thru, Available until, Ends on, Expires. " +
                "Choose the date next to the matching label, never an unrelated date.\n" +
                "Examples:\n" +
                "LABEL: Revised 03/15/2024 | Active Through: March 15, 2027\n" +
                example +
                "LABEL: Date of Revision: January 2023 | Available Until: December 2026\n" +
                "JSON: {\"issue_date\":\"2023-01-01\",\"expiration_date\":\"2026-12-01\",\"products_count\":null}\n" +
                "Configured context phrases to prioritize: " + configuredContexts + "\nSOURCE:\n" + source + "\nJSON:";
        }

        private static PdfTextComparer.PdfCodeInfo? ParseResponse(string output, bool includeProductCount)
        {
            var start = output.IndexOf('{');
            if (start < 0)
                return null;

            var end = output.LastIndexOf('}');
            string jsonText;
            if (end > start)
                jsonText = output.Substring(start, end - start + 1);
            else
                jsonText = output.Substring(start) + "}";

            try
            {
                using var json = JsonDocument.Parse(jsonText);
                var result = new PdfTextComparer.PdfCodeInfo();
                if (json.RootElement.TryGetProperty("latest_code", out var latest) && latest.ValueKind == JsonValueKind.String)
                    result.LatestCode = latest.GetString() ?? "n/a";
                result.IssueDate = FormatDate(json.RootElement, "issue_date");
                result.ExpirationDate = FormatDate(json.RootElement, "expiration_date");
                if (includeProductCount && TryReadInt(json.RootElement, "products_count", out var count))
                    result.ProductsCount = count;
                return result;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static bool TryReadInt(JsonElement root, string property, out int value)
        {
            value = 0;
            if (!root.TryGetProperty(property, out var element))
                return false;

            if (element.ValueKind == JsonValueKind.Number)
                return element.TryGetInt32(out value);

            if (element.ValueKind == JsonValueKind.String &&
                int.TryParse(element.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;

            return false;
        }

        private static string FormatDate(JsonElement root, string property)
        {
            if (!root.TryGetProperty(property, out var value) || value.ValueKind != JsonValueKind.String)
                return "n/a";

            var raw = value.GetString();
            return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
                ? date.ToString("MMM-yyyy", CultureInfo.InvariantCulture)
                : "n/a";
        }

        public void Dispose()
        {
            _gate.Dispose();
            _weights.Dispose();
        }

        private int GetCpuThreadLimit()
        {
            return Math.Max(1, (int)Math.Ceiling(Environment.ProcessorCount * _cpuPercent / 100d));
        }
    }
}
