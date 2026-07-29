using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AIVectorCore
{
    public sealed class SvgGenerationResult
    {
        public string Svg { get; set; } = "";
        public IReadOnlyList<SvgLayer> Layers { get; set; } = new List<SvgLayer>();
        public AiResponse Response { get; set; }
    }

    public sealed class SvgGenerator
    {
        private readonly ApiClient _client;

        public SvgGenerator(ApiClient client)
        {
            _client = client ?? throw new ArgumentNullException(nameof(client));
        }

        public async Task<SvgGenerationResult> GenerateAsync(
            ApiProfile profile,
            GenerationOptions options,
            string prompt,
            ImageInput referenceImage,
            CancellationToken cancellationToken,
            Action<string> onDelta = null)
        {
            options = options ?? new GenerationOptions();
            var hasReference = referenceImage != null && referenceImage.HasValue;
            var isCopyMode = hasReference && string.Equals(options.ReferenceMode, "copy", StringComparison.OrdinalIgnoreCase);
            var lockedTexts = isCopyMode
                ? await TryExtractReferenceTextsAsync(profile, referenceImage, cancellationToken).ConfigureAwait(false)
                : new List<string>();

            var userPrompt = PromptBuilder.BuildUserPrompt(prompt, options.ReferenceMode, hasReference);
            if (lockedTexts.Count > 0) userPrompt = AppendTextLockInstruction(userPrompt, lockedTexts);

            var temperature = options.Temperature;
            if (isCopyMode)
                // 临摹还原更像描摹任务而不是创作任务：继续压低随机性，减少改构图、改人物和改文字。
                temperature = Math.Min(temperature, 0.25);
            var response = await _client.CompleteAsync(
                profile,
                PromptBuilder.BuildSystemPrompt(options),
                userPrompt,
                referenceImage,
                temperature,
                options.MaxTokens,
                null,
                cancellationToken,
                onDelta).ConfigureAwait(false);
            return ToResult(response, lockedTexts);
        }

        public async Task<SvgGenerationResult> EditAsync(
            ApiProfile profile,
            GenerationOptions options,
            string existingSvg,
            string instruction,
            CancellationToken cancellationToken,
            Action<string> onDelta = null)
        {
            options = options ?? new GenerationOptions();
            var system = PromptBuilder.BuildSystemPrompt(options) +
                         "\n\n当前任务是[修改]: 用户会给出现有 SVG 代码和修改要求。请在尽量保持其余内容不变的前提下完成修改, 输出修改后的完整 SVG(仍需满足上述全部要求)。";
            var response = await _client.CompleteAsync(
                profile,
                system,
                PromptBuilder.BuildEditPrompt(existingSvg, instruction),
                null,
                0.4,
                options.MaxTokens,
                null,
                cancellationToken,
                onDelta).ConfigureAwait(false);
            return ToResult(response);
        }

        public async Task<SvgGenerationResult> RefineAgainstReferenceAsync(
            ApiProfile profile,
            GenerationOptions options,
            string existingSvg,
            ImageInput referenceImage,
            CancellationToken cancellationToken,
            Action<string> onDelta = null)
        {
            if (referenceImage == null || !referenceImage.HasValue)
                throw new ArgumentException("请先选择参考图片。", nameof(referenceImage));

            options = options ?? new GenerationOptions();
            var lockedTexts = await TryExtractReferenceTextsAsync(profile, referenceImage, cancellationToken).ConfigureAwait(false);
            var system = PromptBuilder.BuildSystemPrompt(options) +
                         "\n\n当前任务是[对照修正]：用户提供参考图片和当前 SVG。请逐项对比构图、元素、颜色、比例和文字，修正所有差异后只输出完整 SVG。";
            var prompt = "参考图片是目标效果。当前 SVG 代码:\n```svg\n" + (existingSvg ?? "") +
                         "\n```\n\n请根据参考图片修正该 SVG，保持分层并输出完整 SVG。";
            if (lockedTexts.Count > 0) prompt = AppendTextLockInstruction(prompt, lockedTexts);
            var response = await _client.CompleteAsync(
                profile,
                system,
                prompt,
                referenceImage,
                0.25,
                options.MaxTokens,
                null,
                cancellationToken,
                onDelta).ConfigureAwait(false);
            return ToResult(response, lockedTexts);
        }

        private async Task<List<string>> TryExtractReferenceTextsAsync(ApiProfile profile, ImageInput referenceImage, CancellationToken cancellationToken)
        {
            try
            {
                var system = "你只负责 OCR 识别参考图片中的文字。只输出 JSON 数组，例如 [\"文字1\",\"文字2\"]。不要解释，不要猜测，不要补写口号；只列出图片里清晰可见、可逐字确认的文字。";
                var user = "请按从上到下、从左到右的顺序，逐行提取这张参考图中清晰可见的文字。看不清或不确定的文字不要输出。";
                var response = await _client.CompleteAsync(
                    profile,
                    system,
                    user,
                    referenceImage,
                    0.0,
                    1200,
                    null,
                    cancellationToken,
                    null).ConfigureAwait(false);
                return ParseReferenceTexts(response == null ? "" : response.Text);
            }
            catch (OperationCanceledException) { throw; }
            catch
            {
                // OCR 预检失败时不阻塞生成；仍依赖主提示词约束。
                return new List<string>();
            }
        }

        private static string AppendTextLockInstruction(string basePrompt, IReadOnlyList<string> lockedTexts)
        {
            var sb = new StringBuilder(basePrompt ?? "");
            sb.Append("\n\n【参考图文字锁定清单】\n");
            for (var i = 0; i < lockedTexts.Count; i++)
                sb.Append(i + 1).Append(". ").Append(lockedTexts[i]).Append('\n');
            sb.Append("硬性文字规则: SVG 中所有有意义的 <text> 文字只能来自上面的清单, 必须逐字一致; 禁止改写成同义句, 禁止添加口号、说明、小字、宣传语、英文占位或任何清单外文字。若某个文字不在清单中, 宁可不画, 也不要猜。符号或装饰请用 path/line 绘制, 不要用文字代替。");
            return sb.ToString();
        }

        private static List<string> ParseReferenceTexts(string text)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return result;

            foreach (Match m in Regex.Matches(text, "[\\\"“]([^\\\"”]{1,60})[\\\"”]"))
                AddReferenceText(result, seen, m.Groups[1].Value);

            var cleaned = Regex.Replace(text, @"```[a-zA-Z]*", "");
            foreach (var rawLine in Regex.Split(cleaned, @"\r?\n"))
                AddReferenceText(result, seen, rawLine);
            return result;
        }

        private static void AddReferenceText(List<string> result, HashSet<string> seen, string value)
        {
            var s = (value ?? "").Trim();
            if (s.Length == 0) return;
            s = Regex.Replace(s, @"^\s*[-*•\d一二三四五六七八九十]+[\.、:：\)\]\s]+", "");
            s = s.Trim().Trim('"', '\'', '`', '“', '”', '[', ']', ',', '，');
            s = Regex.Replace(s, @"\s+", " ").Trim();
            if (s.Length < 2 || s.Length > 40) return;
            if (Regex.IsMatch(s, @"^(texts?|文字|内容|null|none)$", RegexOptions.IgnoreCase)) return;
            if (Regex.IsMatch(s, @"未识别|没有|无文字|看不清|无法|不确定|清晰可见")) return;
            if (!SvgText.HasSignificantText(s)) return;
            var key = SvgText.NormalizeForCompare(s);
            if (key.Length < 2 || seen.Contains(key)) return;
            seen.Add(key);
            result.Add(s);
        }

        private static SvgGenerationResult ToResult(AiResponse response, IReadOnlyList<string> lockedTexts = null)
        {
            var svg = SvgText.Extract(response == null ? "" : response.Text);
            if (string.IsNullOrWhiteSpace(svg))
                throw new InvalidOperationException("AI 响应中未找到有效 SVG。");
            if (lockedTexts != null && lockedTexts.Count > 0)
                svg = SvgText.LockTextElements(svg, lockedTexts);
            return new SvgGenerationResult
            {
                Svg = svg,
                Layers = SvgLayers.Parse(svg),
                Response = response
            };
        }
    }

    public static class SvgText
    {
        public static string Extract(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var matches = Regex.Matches(text, @"<svg[\s\S]*?</svg>", RegexOptions.IgnoreCase);
            if (matches.Count == 0) return "";
            var svg = matches[matches.Count - 1].Value;
            svg = Regex.Replace(svg, @"<script[\s\S]*?</script>", "", RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg, @"<!--[\s\S]*?-->", "", RegexOptions.IgnoreCase);
            return svg.Trim();
        }

        public static string LockTextElements(string svg, IReadOnlyList<string> allowedTexts)
        {
            if (string.IsNullOrWhiteSpace(svg) || allowedTexts == null || allowedTexts.Count == 0) return svg ?? "";
            var allowed = new List<TextCandidate>();
            foreach (var text in allowedTexts)
            {
                var value = (text ?? "").Trim();
                var norm = NormalizeForCompare(value);
                if (value.Length > 0 && norm.Length > 1) allowed.Add(new TextCandidate { Value = value, Normalized = norm });
            }
            if (allowed.Count == 0) return svg;

            return Regex.Replace(svg, @"<text\b(?<attrs>[^>]*)>(?<content>[\s\S]*?)</text>", match =>
            {
                var rawContent = match.Groups["content"].Value;
                var plain = Regex.Replace(rawContent, @"<[^>]+>", "");
                plain = System.Net.WebUtility.HtmlDecode(plain ?? "").Trim();
                var norm = NormalizeForCompare(plain);
                if (norm.Length == 0 || !HasSignificantText(plain)) return match.Value;

                for (var i = 0; i < allowed.Count; i++)
                {
                    if (string.Equals(norm, allowed[i].Normalized, StringComparison.OrdinalIgnoreCase))
                        return ReplaceTextContent(match.Value, allowed[i].Value);
                }

                var best = default(TextCandidate);
                var bestScore = 0.0;
                for (var i = 0; i < allowed.Count; i++)
                {
                    var score = Similarity(norm, allowed[i].Normalized);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = allowed[i];
                    }
                }

                // 明显是把参考图文字改写/联想了: 保留位置和样式, 但替换回 OCR 原文。
                if (!string.IsNullOrWhiteSpace(best.Value) && bestScore >= 0.45)
                    return ReplaceTextContent(match.Value, best.Value);

                // 清单外长文字直接删除，避免 AI 自己添加宣传语、小字、占位英文。
                return "";
            }, RegexOptions.IgnoreCase);
        }

        public static string NormalizeForCompare(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var decoded = System.Net.WebUtility.HtmlDecode(text) ?? "";
            var sb = new StringBuilder(decoded.Length);
            foreach (var ch in decoded)
            {
                if (char.IsWhiteSpace(ch)) continue;
                if (ch == '·' || ch == '•' || ch == '。' || ch == '，' || ch == ',' || ch == ':' || ch == '：' || ch == '-' || ch == '_' || ch == '—') continue;
                sb.Append(char.ToUpperInvariant(ch));
            }
            return sb.ToString();
        }

        public static bool HasSignificantText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            foreach (var ch in text)
            {
                if (char.IsLetterOrDigit(ch)) return true;
                if (ch >= 0x4e00 && ch <= 0x9fff) return true;
            }
            return false;
        }

        private static string ReplaceTextContent(string fullTextElement, string newText)
        {
            var close = fullTextElement.IndexOf('>');
            if (close < 0) return fullTextElement;
            return fullTextElement.Substring(0, close + 1) + EscapeXmlText(newText) + "</text>";
        }

        private static string EscapeXmlText(string text)
        {
            return (text ?? "")
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static double Similarity(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
            var lcs = LongestCommonSubsequenceLength(a, b);
            return (double)lcs / Math.Max(a.Length, b.Length);
        }

        private static int LongestCommonSubsequenceLength(string a, string b)
        {
            var dp = new int[a.Length + 1, b.Length + 1];
            for (var i = 1; i <= a.Length; i++)
            {
                for (var j = 1; j <= b.Length; j++)
                {
                    if (a[i - 1] == b[j - 1]) dp[i, j] = dp[i - 1, j - 1] + 1;
                    else dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                }
            }
            return dp[a.Length, b.Length];
        }

        private struct TextCandidate
        {
            public string Value;
            public string Normalized;
        }
    }
}
