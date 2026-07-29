using System;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AIVectorCore
{
    public sealed class ImageClient
    {
        private readonly ApiClient _client;

        public ImageClient(ApiClient client)
        {
            _client = client ?? throw new ArgumentNullException(nameof(client));
        }

        public Task<ImageResponse> GenerateAsync(
            ApiProfile profile,
            string model,
            string prompt,
            string resolution,
            CancellationToken cancellationToken)
        {
            var size = string.Equals(resolution, "1k", StringComparison.OrdinalIgnoreCase)
                ? "1024x1024"
                : null;
            var quality = string.IsNullOrWhiteSpace(model) || model.IndexOf("gpt-image", StringComparison.OrdinalIgnoreCase) < 0
                ? null
                : (string.Equals(resolution, "4k", StringComparison.OrdinalIgnoreCase) ? "high"
                   : string.Equals(resolution, "1k", StringComparison.OrdinalIgnoreCase) ? "low" : "medium");
            return _client.GenerateImageAsync(profile, model, prompt, size, quality, cancellationToken);
        }

        public Task<ImageResponse> EditAsync(
            ApiProfile profile,
            string model,
            string prompt,
            string sourceDataUrl,
            string resolution,
            CancellationToken cancellationToken)
        {
            var size = string.Equals(resolution, "1k", StringComparison.OrdinalIgnoreCase)
                ? "1024x1024"
                : null;
            var quality = string.IsNullOrWhiteSpace(model) || model.IndexOf("gpt-image", StringComparison.OrdinalIgnoreCase) < 0
                ? null
                : (string.Equals(resolution, "4k", StringComparison.OrdinalIgnoreCase) ? "high"
                   : string.Equals(resolution, "1k", StringComparison.OrdinalIgnoreCase) ? "low" : "medium");
            return _client.EditImageAsync(profile, model, prompt, sourceDataUrl, size, quality, cancellationToken);
        }

        public async Task<ImageResponse> ChatAsync(
            ApiProfile profile,
            string model,
            string prompt,
            string sourceDataUrl,
            string resolution,
            CancellationToken cancellationToken)
        {
            var chatProfile = profile.Clone();
            chatProfile.Model = string.IsNullOrWhiteSpace(model) ? profile.Model : model;
            var result = await _client.CompleteAsync(
                chatProfile,
                "你是图片生成助手。请直接返回图片数据或图片 URL，不要解释过程。",
                prompt + (string.Equals(resolution, "1k", StringComparison.OrdinalIgnoreCase) ? "" : "\n输出 " + resolution.ToUpperInvariant() + " 分辨率。"),
                string.IsNullOrWhiteSpace(sourceDataUrl) ? null : new ImageInput { DataUrl = sourceDataUrl },
                0.6,
                8000,
                null,
                cancellationToken).ConfigureAwait(false);

            var text = result.Text ?? "";
            var raw = result.RawText ?? "";
            var jsonBase64 = Regex.Match(raw, @"""(?:b64_json|base64|data)""\s*:\s*""([A-Za-z0-9+/=]{128,})""", RegexOptions.IgnoreCase);
            if (jsonBase64.Success) return new ImageResponse { Base64 = jsonBase64.Groups[1].Value };
            var data = Regex.Match(text + "\n" + raw, @"data:image/[a-z0-9.+-]+;base64,([A-Za-z0-9+/=]+)", RegexOptions.IgnoreCase);
            if (data.Success) return new ImageResponse { Base64 = data.Groups[1].Value };
            var jsonUrl = Regex.Match(raw, @"""(?:url|image_url)""\s*:\s*""(https?://[^""\\]+)""", RegexOptions.IgnoreCase);
            if (jsonUrl.Success) return new ImageResponse { Url = jsonUrl.Groups[1].Value };
            var url = Regex.Match(text + "\n" + raw, @"https?://[^\s\]\)""']+", RegexOptions.IgnoreCase);
            if (url.Success) return new ImageResponse { Url = url.Value.TrimEnd('.', ',', ';') };
            throw new InvalidOperationException("对话接口没有返回图片数据或图片 URL。");
        }
    }
}
