using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace AIVectorCore
{
    public sealed class AiResponse
    {
        public string Text { get; set; } = "";
        public int InputTokens { get; set; }
        public int OutputTokens { get; set; }
        public int TotalTokens { get; set; }
        public string RawText { get; set; } = "";
    }

    public sealed class ImageResponse
    {
        public string Base64 { get; set; } = "";
        public string Url { get; set; } = "";

        public bool HasImage
        {
            get { return !string.IsNullOrWhiteSpace(Base64) || !string.IsNullOrWhiteSpace(Url); }
        }
    }

    public sealed class ApiException : Exception
    {
        public HttpStatusCode StatusCode { get; private set; }
        public string ResponseBody { get; private set; }

        public ApiException(string message, HttpStatusCode statusCode, string responseBody)
            : base(message)
        {
            StatusCode = statusCode;
            ResponseBody = responseBody ?? "";
        }
    }

    public sealed class ApiClient : IDisposable
    {
        private readonly HttpClient _http;
        private readonly string _proxy;
        private bool _disposed;

        public ApiClient(string proxy = "", TimeSpan? timeout = null)
        {
            _proxy = proxy ?? "";
            var handler = new HttpClientHandler();
            if (!string.IsNullOrWhiteSpace(proxy))
            {
                handler.Proxy = new WebProxy(proxy);
                handler.UseProxy = true;
            }
            else
            {
                // 空代理时强制直连，避免 CorelDRAW 进程继承失效的系统代理配置。
                handler.UseProxy = false;
            }

            _http = new HttpClient(handler, true);
            _http.Timeout = timeout ?? TimeSpan.FromMinutes(10);
        }

        public async Task<AiResponse> CompleteAsync(
            ApiProfile profile,
            string systemPrompt,
            string userPrompt,
            ImageInput image,
            double temperature,
            int maxTokens,
            IEnumerable<ChatMessage> extraMessages,
            CancellationToken cancellationToken,
            Action<string> onDelta = null)
        {
            EnsureUsable(profile);
            var content = BuildUserContent(profile, userPrompt, image);
            var messages = new List<ChatMessage>();

            if (!profile.IsAnthropic)
                messages.Add(new ChatMessage("system", systemPrompt ?? ""));
            if (extraMessages != null) messages.AddRange(extraMessages);
            messages.Add(new ChatMessage("user", content));

            JObject body;
            string url;
            if (profile.IsAnthropic)
            {
                url = JoinUrl(profile.BaseUrl, "/v1/messages");
                body = new JObject
                {
                    ["model"] = profile.Model,
                    ["max_tokens"] = maxTokens > 0 ? maxTokens : 16000,
                    ["temperature"] = temperature,
                    ["system"] = systemPrompt ?? "",
                    ["messages"] = JToken.FromObject(messages.Select(x => new { role = x.Role, content = x.Content })),
                    ["stream"] = true
                };
            }
            else
            {
                url = JoinUrl(profile.BaseUrl, "/chat/completions");
                body = new JObject
                {
                    ["model"] = profile.Model,
                    ["messages"] = JToken.FromObject(messages.Select(x => new { role = x.Role, content = x.Content })),
                    ["temperature"] = temperature,
                    ["max_tokens"] = maxTokens > 0 ? maxTokens : 16000,
                    ["stream"] = true
                };
            }

            try
            {
                using (var request = CreateJsonRequest(HttpMethod.Post, url, profile, body))
                using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                {
                    return await ReadAiResponseAsync(response, profile.IsAnthropic, onDelta, cancellationToken).ConfigureAwait(false);
                }
            }
            catch (HttpRequestException httpError)
            {
                // 部分本机网络过滤器会阻断 CorelDRAW 进程内的 HttpClient，
                // 但 Windows curl.exe 可正常联网。生成请求也走相同备用通道。
                try
                {
                    return await CompleteViaCurlAsync(profile, url, body, profile.IsAnthropic, cancellationToken, onDelta).ConfigureAwait(false);
                }
                catch (Exception curlError)
                {
                    throw new HttpRequestException(
                        "HTTP 请求失败，curl 备用通道也失败：" + curlError.Message,
                        httpError);
                }
            }
        }

        private async Task<AiResponse> CompleteViaCurlAsync(
            ApiProfile profile,
            string url,
            JObject body,
            bool anthropic,
            CancellationToken cancellationToken,
            Action<string> onDelta)
        {
            var bodyFile = Path.Combine(Path.GetTempPath(), "aivector-request-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(bodyFile, body.ToString(Formatting.None), new UTF8Encoding(false));
                var args = new StringBuilder();
                args.Append("--silent --show-error --fail-with-body --http1.1 -4 --connect-timeout 30 --max-time 600 ");
                args.Append("--noproxy \"*\" ");
                if (!string.IsNullOrWhiteSpace(_proxy))
                    args.Append("--proxy ").Append(QuoteProcessArgument(_proxy)).Append(' ');
                if (profile.IsAnthropic)
                {
                    args.Append("-H ").Append(QuoteProcessArgument("x-api-key: " + (profile.ApiKey ?? ""))).Append(' ');
                    args.Append("-H ").Append(QuoteProcessArgument("anthropic-version: 2023-06-01")).Append(' ');
                }
                else
                {
                    args.Append("-H ").Append(QuoteProcessArgument("Authorization: Bearer " + (profile.ApiKey ?? ""))).Append(' ');
                }
                args.Append("-H ").Append(QuoteProcessArgument("Accept: text/event-stream, application/json")).Append(' ');
                args.Append("-H ").Append(QuoteProcessArgument("Content-Type: application/json")).Append(' ');
                args.Append("--data-binary @").Append(QuoteProcessArgument(bodyFile)).Append(' ');
                args.Append(QuoteProcessArgument(url));

                var startInfo = new ProcessStartInfo
                {
                    FileName = FindCurlExecutable(),
                    Arguments = args.ToString(),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    // curl returns JSON/SSE as UTF-8. Without this override,
                    // .NET Framework uses the Windows code page and corrupts Chinese SVG text.
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
                };

                using (var process = new Process { StartInfo = startInfo })
                {
                    if (!process.Start()) throw new InvalidOperationException("无法启动 curl.exe。");
                    using (cancellationToken.Register(() =>
                    {
                        try { if (!process.HasExited) process.Kill(); } catch { }
                    }))
                    {
                        var stdoutTask = process.StandardOutput.ReadToEndAsync();
                        var stderrTask = process.StandardError.ReadToEndAsync();
                        await Task.Run(() => process.WaitForExit(), cancellationToken).ConfigureAwait(false);
                        var stdout = await stdoutTask.ConfigureAwait(false);
                        var stderr = await stderrTask.ConfigureAwait(false);
                        if (process.ExitCode != 0)
                        {
                            var detail = string.IsNullOrWhiteSpace(stderr) ? "curl.exe 退出码 " + process.ExitCode : stderr.Trim();
                            if (!string.IsNullOrWhiteSpace(stdout)) detail += "\n" + stdout.Substring(0, Math.Min(stdout.Length, 1000));
                            throw new InvalidOperationException(detail);
                        }
                        return ParseCurlAiResponse(stdout, anthropic, onDelta);
                    }
                }
            }
            finally
            {
                try { if (File.Exists(bodyFile)) File.Delete(bodyFile); } catch { }
            }
        }

        private static AiResponse ParseCurlAiResponse(string rawText, bool anthropic, Action<string> onDelta)
        {
            var raw = rawText ?? "";
            if (!System.Text.RegularExpressions.Regex.IsMatch(raw, @"(?:^|\r?\n)data:\s*", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return ParseJsonResponse(raw, anthropic, onDelta);

            var text = new StringBuilder();
            var inputTokens = 0;
            var outputTokens = 0;
            var errorMessage = "";
            using (var reader = new StringReader(raw))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!line.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
                    var payload = line.Substring(5).Trim();
                    if (payload.Length == 0 || payload == "[DONE]") continue;
                    JObject json;
                    try { json = JObject.Parse(payload); }
                    catch (JsonException) { continue; }
                    var error = json["error"];
                    if (error != null) { errorMessage = (string)error["message"] ?? error.ToString(); continue; }
                    var delta = ExtractDelta(json, anthropic, ref inputTokens, ref outputTokens);
                    if (delta.Length > 0)
                    {
                        text.Append(delta);
                        if (onDelta != null) onDelta(delta);
                    }
                }
            }
            if (text.Length == 0 && errorMessage.Length > 0) throw new InvalidDataException(errorMessage);
            return new AiResponse
            {
                Text = text.ToString(),
                RawText = raw,
                InputTokens = inputTokens,
                OutputTokens = outputTokens,
                TotalTokens = inputTokens + outputTokens
            };
        }

        public async Task<IReadOnlyList<string>> FetchModelsAsync(ApiProfile profile, CancellationToken cancellationToken)
        {
            // 获取 /models 时模型名本来就尚未选择，不能使用要求 profile.Model 的完整校验。
            EnsureEndpointUsable(profile);
            try
            {
                return await FetchModelsViaHttpAsync(profile, cancellationToken).ConfigureAwait(false);
            }
            catch (HttpRequestException httpError)
            {
                // CorelDRAW 的宿主进程可能被本机网络过滤器拦截套接字，
                // 但系统 curl.exe 可以正常访问同一地址，作为模型列表和连接测试的备用通道。
                try
                {
                    return await FetchModelsViaCurlAsync(profile, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception curlError)
                {
                    throw new HttpRequestException(
                        "HTTP 请求失败，curl 备用通道也失败：" + curlError.Message,
                        httpError);
                }
            }
        }

        private async Task<IReadOnlyList<string>> FetchModelsViaHttpAsync(
            ApiProfile profile,
            CancellationToken cancellationToken)
        {
            var url = JoinUrl(profile.BaseUrl, "/models");
            using (var request = CreateJsonRequest(HttpMethod.Get, url, profile, null))
            using (var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false))
            {
                var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw CreateApiException(response.StatusCode, text);
                return ParseModelList(text);
            }
        }

        private async Task<IReadOnlyList<string>> FetchModelsViaCurlAsync(
            ApiProfile profile,
            CancellationToken cancellationToken)
        {
            var url = JoinUrl(profile.BaseUrl, "/models");
            var args = new StringBuilder();
            args.Append("--silent --show-error --http1.1 -4 --max-time 60 ");
            args.Append("--noproxy \"*\" ");
            if (!string.IsNullOrWhiteSpace(_proxy))
                args.Append("--proxy ").Append(QuoteProcessArgument(_proxy)).Append(' ');
            if (profile.IsAnthropic)
            {
                args.Append("-H ").Append(QuoteProcessArgument("x-api-key: " + (profile.ApiKey ?? ""))).Append(' ');
                args.Append("-H ").Append(QuoteProcessArgument("anthropic-version: 2023-06-01")).Append(' ');
            }
            else
            {
                args.Append("-H ").Append(QuoteProcessArgument("Authorization: Bearer " + (profile.ApiKey ?? ""))).Append(' ');
            }
            args.Append("-H ").Append(QuoteProcessArgument("Accept: application/json")).Append(' ');
            args.Append(QuoteProcessArgument(url));

            var startInfo = new ProcessStartInfo
            {
                FileName = FindCurlExecutable(),
                Arguments = args.ToString(),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            using (var process = new Process { StartInfo = startInfo })
            {
                if (!process.Start())
                    throw new InvalidOperationException("无法启动 curl.exe。");

                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();
                await Task.Run(() => process.WaitForExit(), cancellationToken).ConfigureAwait(false);
                var stdout = await stdoutTask.ConfigureAwait(false);
                var stderr = await stderrTask.ConfigureAwait(false);

                if (process.ExitCode != 0)
                {
                    var detail = string.IsNullOrWhiteSpace(stderr)
                        ? "curl.exe 退出码 " + process.ExitCode
                        : stderr.Trim();
                    throw new InvalidOperationException(detail);
                }
                return ParseModelList(stdout);
            }
        }

        private static IReadOnlyList<string> ParseModelList(string text)
        {
            JObject json;
            try
            {
                json = JObject.Parse(text ?? "");
            }
            catch (Exception ex)
            {
                throw new InvalidDataException("模型列表响应不是有效 JSON：" + ex.Message);
            }

            var error = json["error"];
            if (error != null)
                throw new InvalidDataException((string)error["message"] ?? error.ToString());

            var result = new List<string>();
            var data = json["data"] as JArray;
            if (data != null)
            {
                foreach (var item in data)
                {
                    var id = (string)item["id"];
                    if (!string.IsNullOrWhiteSpace(id)) result.Add(id);
                }
            }
            return result;
        }

        private static string FindCurlExecutable()
        {
            var windir = Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\Windows";
            var systemCurl = Path.Combine(windir, "System32", "curl.exe");
            return File.Exists(systemCurl) ? systemCurl : "curl.exe";
        }

        private static string QuoteProcessArgument(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }

        public async Task<ImageResponse> GenerateImageAsync(
            ApiProfile profile,
            string model,
            string prompt,
            string size,
            string quality,
            CancellationToken cancellationToken)
        {
            EnsureEndpointUsable(profile);
            var selectedModel = ResolveModel(profile, model);
            var body = new JObject
            {
                ["model"] = selectedModel,
                ["prompt"] = prompt ?? "",
                ["n"] = 1
            };
            if (!string.IsNullOrWhiteSpace(size)) body["size"] = size;
            if (!string.IsNullOrWhiteSpace(quality)) body["quality"] = quality;

            var url = JoinUrl(profile.BaseUrl, "/images/generations");
            try
            {
                using (var request = CreateJsonRequest(HttpMethod.Post, url, profile, body))
                using (var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw CreateApiException(response.StatusCode, text);
                    return ParseImageResponse(text);
                }
            }
            catch (HttpRequestException httpError)
            {
                // CorelDRAW 进程内的 HttpClient 可能被本机网络过滤器拦截，
                // 图像接口和对话接口一样，改用系统 curl.exe 重试同一条路由。
                try
                {
                    return await GenerateImageViaCurlAsync(profile, url, body, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception curlError)
                {
                    throw new HttpRequestException(
                        "图像接口请求失败，curl 备用通道也失败：" + curlError.Message,
                        httpError);
                }
            }
        }

        public async Task<ImageResponse> EditImageAsync(
            ApiProfile profile,
            string model,
            string prompt,
            string dataUrl,
            string size,
            string quality,
            CancellationToken cancellationToken)
        {
            EnsureEndpointUsable(profile);
            var selectedModel = ResolveModel(profile, model);
            var image = DecodeDataUrl(dataUrl);
            var url = JoinUrl(profile.BaseUrl, "/images/edits");
            try
            {
                using (var form = new MultipartFormDataContent())
                {
                    form.Add(new StringContent(selectedModel), "model");
                    form.Add(new StringContent(prompt ?? ""), "prompt");
                    if (!string.IsNullOrWhiteSpace(size)) form.Add(new StringContent(size), "size");
                    if (!string.IsNullOrWhiteSpace(quality)) form.Add(new StringContent(quality), "quality");

                    var file = new ByteArrayContent(image.Bytes);
                    file.Headers.ContentType = new MediaTypeHeaderValue(image.MediaType);
                    // 与可正常工作的 HTA 版本保持一致：中转站的编辑接口使用单数 image。
                    form.Add(file, "image", "source." + image.Extension);

                    using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                    {
                        ApplyAuth(request, profile);
                        request.Content = form;
                        using (var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                        {
                            var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            if (!response.IsSuccessStatusCode) throw CreateApiException(response.StatusCode, text);
                            return ParseImageResponse(text);
                        }
                    }
                }
            }
            catch (HttpRequestException httpError)
            {
                try
                {
                    // HTA 版本使用 MSXML2.ServerXMLHTTP；CorelDRAW 进程内的 HttpClient
                    // 可能被网络过滤器拦截，因此先复用同一套 MSXML multipart 请求。
                    return await EditImageViaMsxmlAsync(profile, url, image, model, prompt, size, quality, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception msxmlError)
                {
                    try
                    {
                        return await EditImageViaCurlAsync(profile, url, image, model, prompt, size, quality, cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception curlError)
                    {
                        throw new HttpRequestException(
                            "图像编辑接口请求失败，HTA/MSXML 与 curl 备用通道均失败："
                            + msxmlError.Message + "；" + curlError.Message,
                            httpError);
                    }
                }
            }
        }

        private async Task<ImageResponse> EditImageViaMsxmlAsync(
            ApiProfile profile,
            string url,
            DecodedImage image,
            string model,
            string prompt,
            string size,
            string quality,
            CancellationToken cancellationToken)
        {
            return await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                var boundary = "----AIVH" + DateTime.Now.Ticks.ToString();
                var body = BuildMultipartBody(boundary, image, model, prompt, size, quality, profile.Model);
                var xhrType = Type.GetTypeFromProgID("MSXML2.ServerXMLHTTP.6.0");
                if (xhrType == null) throw new InvalidOperationException("系统未安装 MSXML2.ServerXMLHTTP.6.0。");

                object xhr = null;
                try
                {
                    xhr = Activator.CreateInstance(xhrType);
                    dynamic request = xhr;
                    request.open("POST", url, false);
                    try { request.setTimeouts(30000, 30000, 60000, 600000); } catch { }
                    if (!string.IsNullOrWhiteSpace(_proxy))
                    {
                        try { request.setProxy(2, _proxy, ""); } catch { }
                    }
                    request.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary);
                    if (profile.IsAnthropic)
                    {
                        request.setRequestHeader("x-api-key", profile.ApiKey ?? "");
                        request.setRequestHeader("anthropic-version", "2023-06-01");
                    }
                    else
                    {
                        request.setRequestHeader("Authorization", "Bearer " + (profile.ApiKey ?? ""));
                    }
                    request.send(body);
                    cancellationToken.ThrowIfCancellationRequested();
                    var status = (HttpStatusCode)Convert.ToInt32(request.status);
                    var text = Convert.ToString(request.responseText) ?? "";
                    if ((int)status < 200 || (int)status >= 300)
                        throw CreateApiException(status, text);
                    return ParseImageResponse(text);
                }
                finally
                {
                    try
                    {
                        if (xhr != null && System.Runtime.InteropServices.Marshal.IsComObject(xhr))
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xhr);
                    }
                    catch { }
                }
            }, cancellationToken).ConfigureAwait(false);
        }

        private static byte[] BuildMultipartBody(
            string boundary,
            DecodedImage image,
            string model,
            string prompt,
            string size,
            string quality,
            string fallbackModel)
        {
            using (var stream = new MemoryStream())
            {
                Action<string> writeText = text =>
                {
                    var bytes = new UTF8Encoding(false).GetBytes(text ?? "");
                    stream.Write(bytes, 0, bytes.Length);
                };

                writeText("--" + boundary + "\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\n"
                    + (string.IsNullOrWhiteSpace(model) ? fallbackModel : model) + "\r\n");
                writeText("--" + boundary + "\r\nContent-Disposition: form-data; name=\"prompt\"\r\n\r\n"
                    + (prompt ?? "") + "\r\n");
                if (!string.IsNullOrWhiteSpace(size))
                    writeText("--" + boundary + "\r\nContent-Disposition: form-data; name=\"size\"\r\n\r\n" + size + "\r\n");
                if (!string.IsNullOrWhiteSpace(quality))
                    writeText("--" + boundary + "\r\nContent-Disposition: form-data; name=\"quality\"\r\n\r\n" + quality + "\r\n");
                writeText("--" + boundary + "\r\nContent-Disposition: form-data; name=\"image\"; filename=\"source."
                    + image.Extension + "\"\r\nContent-Type: " + image.MediaType + "\r\n\r\n");
                stream.Write(image.Bytes, 0, image.Bytes.Length);
                writeText("\r\n--" + boundary + "--\r\n");
                return stream.ToArray();
            }
        }

        private async Task<ImageResponse> GenerateImageViaCurlAsync(
            ApiProfile profile,
            string url,
            JObject body,
            CancellationToken cancellationToken)
        {
            var bodyFile = Path.Combine(Path.GetTempPath(), "aivector-image-request-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(bodyFile, body.ToString(Formatting.None), new UTF8Encoding(false));
                var args = BuildCurlCommonArguments(profile);
                args.Append("-H ").Append(QuoteProcessArgument("Accept: application/json")).Append(' ');
                args.Append("-H ").Append(QuoteProcessArgument("Content-Type: application/json")).Append(' ');
                args.Append("--data-binary @").Append(QuoteProcessArgument(bodyFile)).Append(' ');
                args.Append(QuoteProcessArgument(url));
                return ParseImageResponse(await RunCurlAsync(args.ToString(), cancellationToken).ConfigureAwait(false));
            }
            finally
            {
                try { if (File.Exists(bodyFile)) File.Delete(bodyFile); } catch { }
            }
        }

        private async Task<ImageResponse> EditImageViaCurlAsync(
            ApiProfile profile,
            string url,
            DecodedImage image,
            string model,
            string prompt,
            string size,
            string quality,
            CancellationToken cancellationToken)
        {
            var imageFile = Path.Combine(Path.GetTempPath(), "aivector-image-source-" + Guid.NewGuid().ToString("N") + "." + image.Extension);
            try
            {
                File.WriteAllBytes(imageFile, image.Bytes);
                var args = BuildCurlCommonArguments(profile);
                args.Append("--header ").Append(QuoteProcessArgument("Accept: application/json")).Append(' ');
                args.Append("--form-string ").Append(QuoteProcessArgument("model=" + (string.IsNullOrWhiteSpace(model) ? profile.Model : model))).Append(' ');
                args.Append("--form-string ").Append(QuoteProcessArgument("prompt=" + (prompt ?? ""))).Append(' ');
                if (!string.IsNullOrWhiteSpace(size))
                    args.Append("--form-string ").Append(QuoteProcessArgument("size=" + size)).Append(' ');
                if (!string.IsNullOrWhiteSpace(quality))
                    args.Append("--form-string ").Append(QuoteProcessArgument("quality=" + quality)).Append(' ');
                args.Append("-F ").Append(QuoteProcessArgument("image=@" + imageFile + ";type=" + image.MediaType)).Append(' ');
                args.Append(QuoteProcessArgument(url));
                return ParseImageResponse(await RunCurlAsync(args.ToString(), cancellationToken).ConfigureAwait(false));
            }
            finally
            {
                try { if (File.Exists(imageFile)) File.Delete(imageFile); } catch { }
            }
        }

        private StringBuilder BuildCurlCommonArguments(ApiProfile profile)
        {
            var args = new StringBuilder();
            args.Append("--silent --show-error --fail-with-body --http1.1 -4 --connect-timeout 30 --max-time 600 ");
            args.Append("--noproxy \"*\" ");
            if (!string.IsNullOrWhiteSpace(_proxy))
                args.Append("--proxy ").Append(QuoteProcessArgument(_proxy)).Append(' ');
            if (profile.IsAnthropic)
            {
                args.Append("-H ").Append(QuoteProcessArgument("x-api-key: " + (profile.ApiKey ?? ""))).Append(' ');
                args.Append("-H ").Append(QuoteProcessArgument("anthropic-version: 2023-06-01")).Append(' ');
            }
            else
            {
                args.Append("-H ").Append(QuoteProcessArgument("Authorization: Bearer " + (profile.ApiKey ?? ""))).Append(' ');
            }
            return args;
        }

        private async Task<string> RunCurlAsync(string arguments, CancellationToken cancellationToken)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = FindCurlExecutable(),
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            using (var process = new Process { StartInfo = startInfo })
            {
                if (!process.Start()) throw new InvalidOperationException("无法启动 curl.exe。");
                using (cancellationToken.Register(() =>
                {
                    try { if (!process.HasExited) process.Kill(); } catch { }
                }))
                {
                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();
                    await Task.Run(() => process.WaitForExit(), cancellationToken).ConfigureAwait(false);
                    var stdout = await stdoutTask.ConfigureAwait(false);
                    var stderr = await stderrTask.ConfigureAwait(false);
                    if (process.ExitCode != 0)
                    {
                        var detail = string.IsNullOrWhiteSpace(stderr)
                            ? "curl.exe 退出码 " + process.ExitCode
                            : stderr.Trim();
                        if (!string.IsNullOrWhiteSpace(stdout))
                            detail += "\n" + stdout.Substring(0, Math.Min(stdout.Length, 1000));
                        throw new InvalidOperationException(detail);
                    }
                    return stdout;
                }
            }
        }

        private async Task<AiResponse> ReadAiResponseAsync(
            HttpResponseMessage response,
            bool anthropic,
            Action<string> onDelta,
            CancellationToken cancellationToken)
        {
            var raw = new StringBuilder();
            var text = new StringBuilder();
            var inputTokens = 0;
            var outputTokens = 0;
            var sawSse = false;
            var errorMessage = "";

            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                string line;
                while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    raw.AppendLine(line);
                    if (!line.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
                    sawSse = true;
                    var payload = line.Substring(5).Trim();
                    if (payload.Length == 0 || payload == "[DONE]") continue;

                    JObject json;
                    try { json = JObject.Parse(payload); }
                    catch (JsonException) { continue; }

                    var error = json["error"];
                    if (error != null)
                    {
                        errorMessage = (string)error["message"] ?? error.ToString();
                        continue;
                    }

                    var delta = ExtractDelta(json, anthropic, ref inputTokens, ref outputTokens);
                    if (delta.Length > 0)
                    {
                        text.Append(delta);
                        if (onDelta != null) onDelta(delta);
                    }
                }
            }

            var rawText = raw.ToString();
            if (!response.IsSuccessStatusCode)
                throw CreateApiException(response.StatusCode, rawText);
            if (!sawSse)
                return ParseJsonResponse(rawText, anthropic, onDelta);
            if (text.Length == 0 && errorMessage.Length > 0)
                throw new ApiException(errorMessage, response.StatusCode, rawText);

            return new AiResponse
            {
                Text = text.ToString(),
                RawText = rawText,
                InputTokens = inputTokens,
                OutputTokens = outputTokens,
                TotalTokens = inputTokens + outputTokens
            };
        }

        private static string ExtractDelta(JObject json, bool anthropic, ref int inputTokens, ref int outputTokens)
        {
            if (anthropic)
            {
                var type = (string)json["type"];
                if (type == "content_block_delta")
                    return (string)json["delta"]["text"] ?? "";
                if (type == "message_start")
                    inputTokens += (int?)json["message"]["usage"]["input_tokens"] ?? 0;
                else if (type == "message_delta")
                    outputTokens += (int?)json["usage"]["output_tokens"] ?? 0;
                return "";
            }

            var choices = json["choices"] as JArray;
            var delta = choices != null && choices.Count > 0
                ? choices[0]?["delta"]?["content"]
                : null;
            if (delta != null && delta.Type == JTokenType.String) return (string)delta;
            var usage = json["usage"];
            if (usage != null)
            {
                inputTokens = Math.Max(inputTokens, (int?)usage["prompt_tokens"] ?? 0);
                outputTokens = Math.Max(outputTokens, (int?)usage["completion_tokens"] ?? 0);
            }
            return "";
        }

        private static AiResponse ParseJsonResponse(string rawText, bool anthropic, Action<string> onDelta)
        {
            JObject json;
            try { json = JObject.Parse(rawText); }
            catch (Exception ex) { throw new InvalidDataException("AI 响应不是有效 JSON/SSE: " + ex.Message); }
            var error = json["error"];
            if (error != null) throw new InvalidDataException((string)error["message"] ?? error.ToString());

            var text = new StringBuilder();
            if (anthropic)
            {
                foreach (var item in json["content"] ?? new JArray())
                    if ((string)item["type"] == "text") text.Append((string)item["text"] ?? "");
            }
            else
            {
                var choices = json["choices"] as JArray;
                if (choices != null && choices.Count > 0)
                    text.Append((string)choices[0]?["message"]?["content"] ?? "");
            }
            var result = new AiResponse
            {
                Text = text.ToString(),
                RawText = rawText,
                InputTokens = (int?)json["usage"]?["input_tokens"] ?? (int?)json["usage"]?["prompt_tokens"] ?? 0,
                OutputTokens = (int?)json["usage"]?["output_tokens"] ?? (int?)json["usage"]?["completion_tokens"] ?? 0
            };
            result.TotalTokens = result.InputTokens + result.OutputTokens;
            if (onDelta != null && result.Text.Length > 0) onDelta(result.Text);
            return result;
        }

        private static object BuildUserContent(ApiProfile profile, string userPrompt, ImageInput image)
        {
            var prompt = string.IsNullOrWhiteSpace(userPrompt) ? "请根据参考图片创作。" : userPrompt;
            if (image == null || !image.HasValue) return prompt;
            if (profile.IsAnthropic)
            {
                var parsed = DecodeDataUrl(image.DataUrl);
                return new object[]
                {
                    new { type = "image", source = new { type = "base64", media_type = parsed.MediaType, data = parsed.Base64 } },
                    new { type = "text", text = prompt }
                };
            }
            return new object[]
            {
                new { type = "image_url", image_url = new { url = image.DataUrl } },
                new { type = "text", text = prompt }
            };
        }

        private HttpRequestMessage CreateJsonRequest(HttpMethod method, string url, ApiProfile profile, JObject body)
        {
            var request = new HttpRequestMessage(method, url);
            ApplyAuth(request, profile);
            if (body != null)
            {
                request.Content = new StringContent(body.ToString(), Encoding.UTF8, "application/json");
            }
            return request;
        }

        private static void ApplyAuth(HttpRequestMessage request, ApiProfile profile)
        {
            if (profile.IsAnthropic)
            {
                request.Headers.TryAddWithoutValidation("x-api-key", profile.ApiKey ?? "");
                request.Headers.TryAddWithoutValidation("anthropic-version", "2023-06-01");
            }
            else
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", profile.ApiKey ?? "");
            }
        }

        private static void ValidateProfile(ApiProfile profile)
        {
            ValidateEndpoint(profile);
            if (string.IsNullOrWhiteSpace(profile.Model)) throw new ArgumentException("模型档案缺少模型名。", nameof(profile));
        }

        private static void ValidateEndpoint(ApiProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.BaseUrl)) throw new ArgumentException("模型档案缺少 API 地址。", nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.ApiKey)) throw new ArgumentException("模型档案缺少 API Key。", nameof(profile));
        }

        private static string ResolveModel(ApiProfile profile, string model)
        {
            var selectedModel = string.IsNullOrWhiteSpace(model) ? profile.Model : model;
            if (string.IsNullOrWhiteSpace(selectedModel))
                throw new ArgumentException("生图模型名为空，请先选择图像模型。", nameof(model));
            return selectedModel;
        }

        private static string JoinUrl(string baseUrl, string suffix)
        {
            return (baseUrl ?? "").TrimEnd('/') + "/" + suffix.TrimStart('/');
        }

        private static ApiException CreateApiException(HttpStatusCode status, string body)
        {
            var message = "HTTP " + (int)status;
            try
            {
                var json = JObject.Parse(body ?? "");
                message += " — " + ((string)json["error"]?["message"] ?? json["message"] ?? body);
            }
            catch
            {
                if (!string.IsNullOrWhiteSpace(body)) message += " — " + body.Substring(0, Math.Min(body.Length, 500));
            }
            return new ApiException(message, status, body);
        }

        private static ImageResponse ParseImageResponse(string text)
        {
            var json = JObject.Parse(text);
            var error = json["error"];
            if (error != null) throw new InvalidDataException((string)error["message"] ?? error.ToString());
            var data = json["data"]?[0];
            if (data == null) throw new InvalidDataException("接口未返回图片数据。");
            return new ImageResponse
            {
                Base64 = (string)data["b64_json"] ?? "",
                Url = (string)data["url"] ?? ""
            };
        }

        private sealed class DecodedImage
        {
            public string MediaType;
            public string Base64;
            public byte[] Bytes;
            public string Extension;
        }

        private static DecodedImage DecodeDataUrl(string dataUrl)
        {
            var match = System.Text.RegularExpressions.Regex.Match(
                dataUrl ?? "", @"^data:(image/[a-z0-9.+-]+);base64,(.*)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!match.Success) throw new ArgumentException("图片必须是 data:image/...;base64,... 格式。", nameof(dataUrl));
            var mediaType = match.Groups[1].Value.ToLowerInvariant();
            var ext = mediaType.Substring("image/".Length);
            if (ext == "jpeg") ext = "jpg";
            return new DecodedImage
            {
                MediaType = mediaType,
                Base64 = match.Groups[2].Value,
                Bytes = Convert.FromBase64String(match.Groups[2].Value),
                Extension = ext
            };
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ApiClient));
        }

        private void EnsureUsable(ApiProfile profile)
        {
            EnsureNotDisposed();
            ValidateProfile(profile);
        }

        private void EnsureEndpointUsable(ApiProfile profile)
        {
            EnsureNotDisposed();
            ValidateEndpoint(profile);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _http.Dispose();
        }
    }
}
